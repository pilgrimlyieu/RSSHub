import { once } from 'node:events';
import type { RequestListener, Server } from 'node:http';
import { createServer } from 'node:http';
import type { Socket } from 'node:net';
import { inspect } from 'node:util';

import { afterEach, beforeEach, describe, expect, it, vi } from 'vitest';

import type { Config } from '../lib/config';
import type { OutlookMailAccount } from '../lib/utils/mail-auth/account';
import { MailAuthError } from '../lib/utils/mail-auth/errors';
import { createOutlookOAuth } from '../lib/utils/mail-auth/outlook';
import type { OAuthRequest, OAuthResponse } from '../lib/utils/mail-auth/transport';
import { postOAuthForm } from '../lib/utils/mail-auth/transport';

const mocks = vi.hoisted(() => ({
    config: {
        proxy: { strategy: 'all', url_regex: '.*' },
    } as Pick<Config, 'proxy' | 'proxyUri' | 'proxyUris' | 'pacUri' | 'pacScript'>,
}));

vi.mock('../lib/config', () => ({ config: mocks.config }));
vi.mock('../lib/utils/proxy', () => {
    throw new Error('Mail OAuth must not import the global proxy dispatcher.');
});
vi.mock('../lib/utils/request-rewriter', () => {
    throw new Error('Mail OAuth must not import the global request rewriter.');
});
vi.mock('../lib/utils/ofetch', () => {
    throw new Error('Mail OAuth must not import the retrying request wrapper.');
});
vi.mock('../lib/utils/logger', () => {
    throw new Error('Mail OAuth transport must not import logging proxy normalizers.');
});

const account: OutlookMailAccount = {
    email: 'alice@outlook.com',
    username: 'alice@outlook.com',
    host: 'outlook.office365.com',
    port: 993,
    auth: 'oauth2',
    provider: 'outlook',
    clientId: '00000000-0000-4000-8000-000000000000',
    tenant: 'consumers',
};
const accessToken = 'private-access-token';
const refreshToken = 'private-refresh-token+with&form=characters';
const deviceCredential = 'private-device-credential';
const upstreamDetail = 'private-upstream-error-description';
const scopes = 'https://outlook.office.com/IMAP.AccessAsUser.All offline_access';
const servers: Server[] = [];
const sockets = new Set<Socket>();

const deviceResponse = (overrides?: Record<string, unknown>): OAuthResponse => ({
    status: 200,
    body: { device_code: deviceCredential, user_code: 'ABCD-EFGH', verification_uri: 'https://microsoft.com/devicelogin', expires_in: 60, interval: 2, ...overrides },
});

const tokenResponse = (overrides?: Record<string, unknown>): OAuthResponse => ({
    status: 200,
    body: { access_token: accessToken, refresh_token: refreshToken, expires_in: 3600, token_type: 'Bearer', ...overrides },
});

const oauthError = (error: string, status = 400): OAuthResponse => ({ status, body: { error, error_description: `${upstreamDetail}: ${accessToken} ${refreshToken} ${deviceCredential}` } });

const createProtocol = () => {
    const clock = { now: 1_800_000_000_000 };
    const request = vi.fn<OAuthRequest>();
    const wait = vi.fn((milliseconds: number) => {
        clock.now += milliseconds;
        return Promise.resolve();
    });
    return { clock, request, wait, ...createOutlookOAuth({ request, now: () => clock.now, wait }) };
};

const expectRedacted = async (promise: Promise<unknown>, message: RegExp) => {
    let actualError: unknown;
    try {
        await promise;
    } catch (error) {
        actualError = error;
    }
    expect(actualError).toBeInstanceOf(MailAuthError);
    expect((actualError as Error).message).toMatch(message);
    const details = inspect(actualError, { showHidden: true, depth: 8 });
    for (const secret of [accessToken, refreshToken, deviceCredential, upstreamDetail]) {
        expect(details).not.toContain(secret);
    }
    for (const field of ['cause', 'response', 'data']) {
        expect(actualError).not.toHaveProperty(field);
    }
};

const startServer = async (handler: RequestListener) => {
    const server = createServer(handler);
    server.on('connection', (socket) => {
        sockets.add(socket);
        socket.once('close', () => sockets.delete(socket));
    });
    servers.push(server);
    server.listen(0, '127.0.0.1');
    await once(server, 'listening');
    const address = server.address();
    if (!address || typeof address === 'string') {
        throw new Error('The OAuth test server did not open a TCP port.');
    }
    return { server, url: `http://127.0.0.1:${address.port}` };
};

const closeServer = (server: Server) => new Promise<void>((resolve, reject) => server.close((error) => (error ? reject(error) : resolve())));

beforeEach(() => {
    mocks.config.proxy = { strategy: 'all', url_regex: '.*' };
    delete mocks.config.proxyUri;
    delete mocks.config.proxyUris;
    delete mocks.config.pacUri;
    delete mocks.config.pacScript;
});

afterEach(async () => {
    for (const socket of sockets) {
        socket.destroy();
    }
    await Promise.all(servers.map((server) => closeServer(server)));
    servers.length = 0;
    vi.restoreAllMocks();
    vi.unstubAllGlobals();
});

describe('Outlook device authorization', () => {
    it('uses fixed official endpoints and scopes, waits in order, and retains cumulative slow_down delays', async () => {
        const api = createProtocol();
        api.request
            .mockResolvedValueOnce(deviceResponse())
            .mockResolvedValueOnce(oauthError('authorization_pending'))
            .mockResolvedValueOnce(oauthError('slow_down'))
            .mockResolvedValueOnce(oauthError('slow_down'))
            .mockResolvedValueOnce(oauthError('authorization_pending'))
            .mockResolvedValueOnce(tokenResponse());
        const onDeviceCode = vi.fn();

        const result = await api.authorizeOutlook(account, onDeviceCode);

        expect(onDeviceCode).toHaveBeenCalledExactlyOnceWith({ verificationUri: 'https://microsoft.com/devicelogin', userCode: 'ABCD-EFGH', expiresIn: 60 });
        expect(api.wait.mock.calls.map(([milliseconds]) => milliseconds)).toEqual([2000, 2000, 7000, 12000, 12000]);
        expect(api.request.mock.calls[0][0]).toBe('https://login.microsoftonline.com/consumers/oauth2/v2.0/devicecode');
        expect(Object.fromEntries(api.request.mock.calls[0][1])).toEqual({ client_id: account.clientId, scope: scopes });
        for (const [url, form, options] of api.request.mock.calls.slice(1)) {
            expect(url).toBe('https://login.microsoftonline.com/consumers/oauth2/v2.0/token');
            expect(Object.fromEntries(form)).toEqual({ grant_type: 'urn:ietf:params:oauth:grant-type:device_code', client_id: account.clientId, device_code: deviceCredential });
            expect(options?.timeoutMs).toBeLessThanOrEqual(15000);
        }
        expect(result).toEqual({ accessToken, refreshToken, expiresAt: api.clock.now + 3_600_000 });
    });

    it('defaults to the protocol polling interval when interval is omitted', async () => {
        const api = createProtocol();
        api.request.mockResolvedValueOnce(deviceResponse({ interval: undefined })).mockResolvedValueOnce(tokenResponse());
        await api.authorizeOutlook(account, vi.fn());
        expect(api.wait).toHaveBeenCalledExactlyOnceWith(5000);
    });

    it.each(['authorization_declined', 'access_denied'])('stops immediately after %s and hides the upstream response', async (code) => {
        const api = createProtocol();
        api.request.mockResolvedValueOnce(deviceResponse()).mockResolvedValueOnce(oauthError(code));
        await expectRedacted(api.authorizeOutlook(account, vi.fn()), /denied.*login/i);
        expect(api.request).toHaveBeenCalledTimes(2);
        expect(api.wait).toHaveBeenCalledTimes(1);
    });

    it.each(['invalid_client', 'unauthorized_client', 'invalid_scope'])('explains application configuration errors: %s', async (code) => {
        const api = createProtocol();
        api.request.mockResolvedValueOnce(oauthError(code));
        const prompt = vi.fn();
        await expectRedacted(api.authorizeOutlook(account, prompt), /clientId.*tenant.*public-client/i);
        expect(prompt).not.toHaveBeenCalled();
        expect(api.request).toHaveBeenCalledTimes(1);
    });

    it('stops at the server deadline without sending a request after expiry', async () => {
        const api = createProtocol();
        api.request.mockResolvedValueOnce(deviceResponse({ expires_in: 3, interval: 5 }));
        await expectRedacted(api.authorizeOutlook(account, vi.fn()), /expired.*login/i);
        expect(api.wait).toHaveBeenCalledExactlyOnceWith(3000);
        expect(api.request).toHaveBeenCalledTimes(1);
    });

    it('limits an in-flight poll to the remaining deadline and rejects a late successful response', async () => {
        const api = createProtocol();
        api.request.mockResolvedValueOnce(deviceResponse({ expires_in: 5 })).mockImplementationOnce(() => {
            api.clock.now += 3000;
            return Promise.resolve(tokenResponse());
        });
        await expectRedacted(api.authorizeOutlook(account, vi.fn()), /expired.*login/i);
        expect(api.request.mock.calls[1][2]).toEqual({ timeoutMs: 3000 });
        expect(api.request).toHaveBeenCalledTimes(2);
    });

    it.each(['expired_token', 'bad_verification_code'])('stops when Microsoft rejects the device credential as %s', async (code) => {
        const api = createProtocol();
        api.request.mockResolvedValueOnce(deviceResponse()).mockResolvedValueOnce(oauthError(code));
        await expectRedacted(api.authorizeOutlook(account, vi.fn()), /expired.*login/i);
        expect(api.request).toHaveBeenCalledTimes(2);
    });

    it('does not retry a poll after a network failure or leak its cause', async () => {
        const api = createProtocol();
        api.request.mockResolvedValueOnce(deviceResponse()).mockRejectedValueOnce(new Error(upstreamDetail, { cause: new Error(refreshToken) }));
        await expectRedacted(api.authorizeOutlook(account, vi.fn()), /network\/proxy/i);
        expect(api.request).toHaveBeenCalledTimes(2);
    });

    it.each([
        { device_code: '' },
        { user_code: undefined },
        { verification_uri: 'http://microsoft.com/devicelogin' },
        { verification_uri: 'https://not-microsoft.example/devicelogin' },
        { expires_in: 0 },
        { expires_in: '60' },
        { interval: 0 },
        { interval: null },
        { interval: -1 },
    ])('rejects malformed device responses before displaying a prompt: %j', async (overrides) => {
        const api = createProtocol();
        api.request.mockResolvedValueOnce(deviceResponse(overrides));
        const prompt = vi.fn();
        await expectRedacted(api.authorizeOutlook(account, prompt), /invalid response/i);
        expect(prompt).not.toHaveBeenCalled();
        expect(api.request).toHaveBeenCalledTimes(1);
    });

    it('requires offline access before returning a newly authorized token set', async () => {
        const api = createProtocol();
        api.request.mockResolvedValueOnce(deviceResponse()).mockResolvedValueOnce(tokenResponse({ refresh_token: undefined }));
        await expectRedacted(api.authorizeOutlook(account, vi.fn()), /offline access/i);
    });
});

describe('Outlook token refresh', () => {
    it('sends only the expected form fields, forwards cancellation, and timestamps the received response', async () => {
        const api = createProtocol();
        const controller = new AbortController();
        const startedAt = api.clock.now;
        api.request.mockImplementationOnce(() => {
            api.clock.now += 1750;
            return Promise.resolve(tokenResponse({ refresh_token: 'rotated-refresh-token' }));
        });

        const result = await api.refreshOutlook(account, refreshToken, controller.signal);

        const [url, form, options] = api.request.mock.calls[0];
        expect(url).toBe('https://login.microsoftonline.com/consumers/oauth2/v2.0/token');
        expect(Object.fromEntries(form)).toEqual({ grant_type: 'refresh_token', client_id: account.clientId, refresh_token: refreshToken, scope: scopes });
        expect(options?.signal).toBe(controller.signal);
        expect(result).toEqual({ accessToken, refreshToken: 'rotated-refresh-token', expiresAt: startedAt + 1750 + 3_600_000 });
    });

    it('leaves a missing replacement refresh token undefined so the store can preserve the old one', async () => {
        const api = createProtocol();
        api.request.mockResolvedValueOnce(tokenResponse({ refresh_token: undefined }));
        const result = await api.refreshOutlook(account, refreshToken);
        expect(result).toEqual({ accessToken, expiresAt: api.clock.now + 3_600_000 });
        expect(result).not.toHaveProperty('refreshToken');
    });

    it.each(['invalid_grant', 'interaction_required', 'login_required', 'consent_required'])('asks for reauthorization on %s without echoing server data', async (code) => {
        const api = createProtocol();
        api.request.mockResolvedValueOnce(oauthError(code));
        await expectRedacted(api.refreshOutlook(account, refreshToken), /mail-auth login again/i);
        expect(api.request).toHaveBeenCalledTimes(1);
    });

    it.each([{ access_token: '' }, { access_token: undefined }, { token_type: 'Basic' }, { token_type: undefined }, { expires_in: '3600' }, { expires_in: 0 }, { expires_in: Infinity }, { refresh_token: null }])(
        'rejects malformed token fields without attaching the response: %j',
        async (overrides) => {
            const api = createProtocol();
            api.request.mockResolvedValueOnce(tokenResponse(overrides));
            await expectRedacted(api.refreshOutlook(account, refreshToken), /invalid response/i);
        }
    );

    it('rejects missing refresh credentials without contacting Microsoft', async () => {
        const api = createProtocol();
        await expectRedacted(api.refreshOutlook(account, ''), /login first/i);
        expect(api.request).not.toHaveBeenCalled();
    });
});

describe('Mail OAuth HTTP transport', () => {
    it.each([500, 421])('sends a failed HTTP %s token POST exactly once and hides the response in its public error', async (status) => {
        const requests = vi.fn();
        const { url } = await startServer((request, response) => {
            requests(request.method, request.headers['content-type']);
            response.writeHead(status, { 'content-type': 'application/json' });
            response.end(JSON.stringify(oauthError('upstream_error', status).body));
        });
        const globalFetch = vi.fn().mockRejectedValue(new Error('The global fetch wrapper must not be used.'));
        vi.stubGlobal('fetch', globalFetch);
        const api = createOutlookOAuth({ request: (_url, form, options) => postOAuthForm(url, form, options), now: Date.now, wait: () => Promise.resolve() });

        await expectRedacted(api.refreshOutlook(account, refreshToken), status === 500 ? /temporarily unavailable/i : /rejected/i);
        expect(requests).toHaveBeenCalledExactlyOnceWith('POST', 'application/x-www-form-urlencoded');
        expect(globalFetch).not.toHaveBeenCalled();
    });

    it('does not retry a disconnected POST', async () => {
        const requests = vi.fn();
        const { url } = await startServer((request) => {
            requests(request.method);
            request.socket.destroy();
        });
        await expectRedacted(postOAuthForm(url, new URLSearchParams({ refresh_token: refreshToken })), /request failed/i);
        expect(requests).toHaveBeenCalledExactlyOnceWith('POST');
    });

    it('rejects a 307 without forwarding token data to the redirect target', async () => {
        const targetRequests = vi.fn();
        const target = await startServer((_request, response) => {
            targetRequests();
            response.end('{}');
        });
        const originRequests = vi.fn();
        const origin = await startServer((_request, response) => {
            originRequests();
            response.writeHead(307, { location: `${target.url}/token` });
            response.end();
        });
        await expectRedacted(postOAuthForm(`${origin.url}/token`, new URLSearchParams({ refresh_token: refreshToken })), /request failed/i);
        expect(originRequests).toHaveBeenCalledTimes(1);
        expect(targetRequests).not.toHaveBeenCalled();
    });

    it('sanitizes JSON parse failures and applies its timeout until the response body finishes', async () => {
        const { url } = await startServer((request, response) => {
            response.writeHead(200, { 'content-type': 'application/json' });
            if (request.url === '/incomplete') {
                response.write(`{"access_token":"${accessToken}`);
                return;
            }
            response.end(`${upstreamDetail} ${refreshToken}`);
        });
        await expectRedacted(postOAuthForm(`${url}/invalid`, new URLSearchParams()), /invalid response/i);
        await expectRedacted(postOAuthForm(`${url}/incomplete`, new URLSearchParams(), { timeoutMs: 50 }), /timed out/i);
    });

    it('aborts an in-flight request without exposing an abort reason', async () => {
        const controller = new AbortController();
        const { url } = await startServer(() => controller.abort(new Error(`${refreshToken} ${upstreamDetail}`)));
        await expectRedacted(postOAuthForm(url, new URLSearchParams({ refresh_token: refreshToken }), { signal: controller.signal }), /cancelled/i);
    });

    it.each(['on_retry', 'unmatched'])('uses a direct connection for %s without importing unsupported proxy agents', async (mode) => {
        const requests = vi.fn();
        const { url } = await startServer((_request, response) => {
            requests();
            response.end('{}');
        });
        mocks.config.proxyUri = 'socks5://127.0.0.1:1';
        mocks.config.pacScript = 'unavailable proxy auto-configuration';
        if (mode === 'on_retry') {
            mocks.config.proxy.strategy = 'on_retry';
        } else {
            mocks.config.proxy.url_regex = '^https://login.microsoftonline.com/';
        }
        expect(await postOAuthForm(url, new URLSearchParams())).toEqual({ status: 200, body: {} });
        expect(requests).toHaveBeenCalledTimes(1);
    });

    it.each(['socks', 'pac', 'invalid'])('rejects a required %s proxy instead of silently sending directly', async (kind) => {
        const requests = vi.fn();
        const { url } = await startServer((_request, response) => {
            requests();
            response.end('{}');
        });
        if (kind === 'pac') {
            mocks.config.pacUri = `https://proxy.example/${upstreamDetail}`;
        } else {
            mocks.config.proxyUri = kind === 'socks' ? 'socks5://127.0.0.1:1080' : `http://invalid host/${refreshToken}`;
        }
        await expectRedacted(postOAuthForm(url, new URLSearchParams()), kind === 'invalid' ? /proxy configuration is invalid/i : /HTTP\(S\) proxies only/i);
        expect(requests).not.toHaveBeenCalled();
    });

    it('selects the first configured HTTP proxy once and never falls back after failure', async () => {
        const directRequests = vi.fn();
        const target = await startServer((_request, response) => {
            directRequests();
            response.end('{}');
        });
        const proxyRequests = vi.fn();
        const proxy = await startServer((request, response) => {
            proxyRequests(request.headers['proxy-authorization']);
            response.writeHead(407);
            response.end();
        });
        proxy.server.on('connect', (request, socket) => {
            proxyRequests(request.headers['proxy-authorization']);
            socket.end('HTTP/1.1 407 Proxy Authentication Required\r\nContent-Length: 0\r\nConnection: close\r\n\r\n');
        });
        const alternateRequests = vi.fn();
        const alternate = await startServer((_request, response) => {
            alternateRequests();
            response.end('{}');
        });
        alternate.server.on('connect', (_request, socket) => {
            alternateRequests();
            socket.destroy();
        });
        const proxyAuth = Buffer.from('proxy-user:private-proxy-password').toString('base64');
        mocks.config.proxyUris = [proxy.url, alternate.url];
        mocks.config.proxy.auth = proxyAuth;

        await expectRedacted(postOAuthForm(target.url, new URLSearchParams({ refresh_token: refreshToken })), /request failed/i);
        expect(proxyRequests).toHaveBeenCalledExactlyOnceWith(`Basic ${proxyAuth}`);
        expect(alternateRequests).not.toHaveBeenCalled();
        expect(directRequests).not.toHaveBeenCalled();
    });
});
