import { setTimeout as wait } from 'node:timers/promises';

import type { OutlookMailAccount } from './account';
import { MailAuthError } from './errors';
import type { OAuthRequest, OAuthRequestOptions, OAuthResponse } from './transport';
import { postOAuthForm } from './transport';

export interface DeviceCodePrompt {
    verificationUri: string;
    userCode: string;
    expiresIn: number;
}

export interface OAuthTokenSet {
    accessToken: string;
    refreshToken?: string;
    expiresAt: number;
}

interface OutlookDependencies {
    request: OAuthRequest;
    now: () => number;
    wait: (milliseconds: number) => Promise<void>;
}

const scope = 'https://outlook.office.com/IMAP.AccessAsUser.All offline_access';
const verificationHosts = new Set(['microsoft.com', 'www.microsoft.com', 'login.microsoftonline.com']);
const endpoint = (account: OutlookMailAccount, resource: 'devicecode' | 'token') => `https://login.microsoftonline.com/${encodeURIComponent(account.tenant)}/oauth2/v2.0/${resource}`;
const isRecord = (value: unknown): value is Record<string, unknown> => typeof value === 'object' && value !== null && !Array.isArray(value);
const isNonemptyString = (value: unknown): value is string => typeof value === 'string' && value.trim().length > 0;
const isPositiveInteger = (value: unknown): value is number => typeof value === 'number' && Number.isSafeInteger(value) && value > 0;
const malformedResponse = () => new MailAuthError('Microsoft OAuth returned an invalid response. Retry mail-auth login; if this persists, check the application configuration.');
const expiredAuthorization = () => new MailAuthError('Outlook device authorization expired. Run mail-auth login again and finish signing in before the code expires.');

const requestOAuth = async (request: OAuthRequest, url: string, form: URLSearchParams, options?: OAuthRequestOptions): Promise<OAuthResponse> => {
    try {
        return await request(url, form, options);
    } catch (error) {
        if (error instanceof MailAuthError) {
            throw error;
        }
        throw new MailAuthError('Microsoft OAuth request failed. Check network/proxy settings and retry the mail-auth command.');
    }
};

const expiryTime = (seconds: unknown, receivedAt: number): number => {
    if (!isPositiveInteger(seconds)) {
        throw malformedResponse();
    }
    const expiresAt = receivedAt + seconds * 1000;
    if (!Number.isSafeInteger(expiresAt)) {
        throw malformedResponse();
    }
    return expiresAt;
};

const responseError = (response: OAuthResponse): MailAuthError => {
    const code = isRecord(response.body) ? response.body.error : undefined;
    if (response.status >= 500 || code === 'server_error' || code === 'temporarily_unavailable') {
        return new MailAuthError('Microsoft OAuth is temporarily unavailable. Retry the mail-auth command later.');
    }
    switch (code) {
        case 'authorization_declined':
        case 'access_denied':
            return new MailAuthError('Outlook authorization was denied. Check any organization sign-in policy, then run mail-auth login again when ready.');
        case 'expired_token':
        case 'bad_verification_code':
            return expiredAuthorization();
        case 'invalid_grant':
        case 'interaction_required':
        case 'login_required':
        case 'consent_required':
            return new MailAuthError('Outlook authorization expired or needs user approval. Run mail-auth login again.');
        case 'invalid_client':
        case 'unauthorized_client':
        case 'invalid_scope':
        case 'invalid_request':
            return new MailAuthError('Check the Outlook clientId, tenant, supported account types and public-client permissions, then run mail-auth login again.');
        default:
            return new MailAuthError('Microsoft OAuth rejected the request. Check the application configuration and sign-in policy, then retry mail-auth login.');
    }
};

const parseTokens = (response: OAuthResponse, receivedAt: number): OAuthTokenSet => {
    if (response.status !== 200) {
        throw responseError(response);
    }
    const body = response.body;
    if (!isRecord(body) || !isNonemptyString(body.access_token) || typeof body.token_type !== 'string' || body.token_type.toLowerCase() !== 'bearer') {
        throw malformedResponse();
    }
    if (body.refresh_token !== undefined && !isNonemptyString(body.refresh_token)) {
        throw malformedResponse();
    }
    return {
        accessToken: body.access_token,
        ...(isNonemptyString(body.refresh_token) && { refreshToken: body.refresh_token }),
        expiresAt: expiryTime(body.expires_in, receivedAt),
    };
};

const parseDeviceCode = (response: OAuthResponse, receivedAt: number) => {
    if (response.status !== 200) {
        throw responseError(response);
    }
    const body = response.body;
    if (!isRecord(body) || !isNonemptyString(body.device_code) || !isNonemptyString(body.user_code) || !isNonemptyString(body.verification_uri) || !isPositiveInteger(body.expires_in)) {
        throw malformedResponse();
    }
    const interval = body.interval === undefined ? 5 : body.interval;
    if (!isPositiveInteger(interval) || !Number.isSafeInteger(interval * 1000)) {
        throw malformedResponse();
    }
    try {
        const url = new URL(body.verification_uri);
        if (url.protocol !== 'https:' || !verificationHosts.has(url.hostname) || url.username || url.password) {
            throw malformedResponse();
        }
    } catch {
        throw malformedResponse();
    }
    return {
        deviceCode: body.device_code,
        prompt: { verificationUri: body.verification_uri, userCode: body.user_code, expiresIn: body.expires_in },
        deadline: expiryTime(body.expires_in, receivedAt),
        intervalMs: interval * 1000,
    };
};

export const createOutlookOAuth = (dependencies: OutlookDependencies) => {
    const authorizeOutlook = async (account: OutlookMailAccount, onDeviceCode: (prompt: DeviceCodePrompt) => void): Promise<OAuthTokenSet & { refreshToken: string }> => {
        const response = await requestOAuth(dependencies.request, endpoint(account, 'devicecode'), new URLSearchParams({ client_id: account.clientId, scope }));
        const device = parseDeviceCode(response, dependencies.now());
        try {
            onDeviceCode(device.prompt);
        } catch {
            throw new MailAuthError('Unable to display the Outlook sign-in prompt. Run mail-auth login again.');
        }

        let intervalMs = device.intervalMs;
        while (dependencies.now() < device.deadline) {
            // OAuth device-code polling must be sequential and follow the server's interval.
            // oxlint-disable-next-line no-await-in-loop
            await dependencies.wait(Math.min(intervalMs, device.deadline - dependencies.now()));
            const remainingMs = device.deadline - dependencies.now();
            if (remainingMs <= 0) {
                throw expiredAuthorization();
            }

            let tokenResponse: OAuthResponse;
            try {
                // oxlint-disable-next-line no-await-in-loop -- A pending device grant must not issue concurrent token requests.
                tokenResponse = await requestOAuth(
                    dependencies.request,
                    endpoint(account, 'token'),
                    new URLSearchParams({ grant_type: 'urn:ietf:params:oauth:grant-type:device_code', client_id: account.clientId, device_code: device.deviceCode }),
                    { timeoutMs: Math.min(15000, remainingMs) }
                );
            } catch (error) {
                if (dependencies.now() >= device.deadline) {
                    throw expiredAuthorization();
                }
                throw error;
            }
            const receivedAt = dependencies.now();
            if (receivedAt >= device.deadline) {
                throw expiredAuthorization();
            }

            const code = isRecord(tokenResponse.body) ? tokenResponse.body.error : undefined;
            if (tokenResponse.status === 400 && code === 'authorization_pending') {
                continue;
            }
            if (tokenResponse.status === 400 && code === 'slow_down') {
                intervalMs += 5000;
                continue;
            }
            const tokens = parseTokens(tokenResponse, receivedAt);
            if (!tokens.refreshToken) {
                throw new MailAuthError('Outlook did not grant offline access. Run mail-auth login again and allow offline access.');
            }
            return { ...tokens, refreshToken: tokens.refreshToken };
        }
        throw expiredAuthorization();
    };

    const refreshOutlook = async (account: OutlookMailAccount, refreshToken: string, signal?: AbortSignal): Promise<OAuthTokenSet> => {
        if (!isNonemptyString(refreshToken)) {
            throw new MailAuthError('Outlook has no refresh token. Run mail-auth login first.');
        }
        const response = await requestOAuth(dependencies.request, endpoint(account, 'token'), new URLSearchParams({ grant_type: 'refresh_token', client_id: account.clientId, refresh_token: refreshToken, scope }), { signal });
        return parseTokens(response, dependencies.now());
    };

    return { authorizeOutlook, refreshOutlook };
};

export const { authorizeOutlook, refreshOutlook } = createOutlookOAuth({ request: postOAuthForm, now: Date.now, wait });
