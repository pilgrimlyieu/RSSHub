import { Readable } from 'node:stream';

import type { Dispatcher } from 'undici';
import { Agent, fetch as undiciFetch, ProxyAgent } from 'undici';

import { config } from '@/config';

import { MailAuthError } from './errors';

export interface OAuthResponse {
    status: number;
    body: unknown;
}

export interface OAuthRequestOptions {
    signal?: AbortSignal;
    timeoutMs?: number;
}

export type OAuthRequest = (url: string, form: URLSearchParams, options?: OAuthRequestOptions) => Promise<OAuthResponse>;

const requestTimeoutMs = 15000;
const directDispatcher = new Agent();
const proxyDispatchers = new Map<string, ProxyAgent>();

const unsupportedProxy = () => new MailAuthError('Mail OAuth supports HTTP(S) proxies only. Replace SOCKS/PAC with an HTTP(S) proxy for matching OAuth requests.');
const invalidProxy = () => new MailAuthError('Mail OAuth proxy configuration is invalid. Check PROXY_URI / PROXY_URIS or PROXY_HOST, PROXY_PORT and PROXY_AUTH.');

const getProxyDispatcher = (url: string): Dispatcher => {
    if (config.proxy.strategy === 'on_retry') {
        return directDispatcher;
    }
    if (config.proxy.strategy !== 'all') {
        throw new MailAuthError('Set PROXY_STRATEGY to all to proxy Mail OAuth, or on_retry to use a direct connection without retries.');
    }

    let matches: boolean;
    try {
        matches = new RegExp(config.proxy.url_regex).test(url);
    } catch {
        throw new MailAuthError('PROXY_URL_REGEX is invalid. Correct the expression before authorizing mail.');
    }
    if (!matches) {
        return directDispatcher;
    }
    if (config.pacUri || config.pacScript) {
        throw unsupportedProxy();
    }

    // A token POST selects one proxy and never falls through to another proxy or a direct connection.
    const proxyUri = config.proxyUris?.[0] ?? config.proxyUri;
    const proxyHost = proxyUri || config.proxy.host;
    if (!proxyHost) {
        if (config.proxy.protocol || config.proxy.port || config.proxy.auth) {
            throw invalidProxy();
        }
        return directDispatcher;
    }

    let proxyUrl: URL;
    try {
        const protocol = proxyUri ? 'http' : (config.proxy.protocol ?? 'http');
        proxyUrl = new URL(proxyHost.includes('://') ? proxyHost : `${protocol}://${proxyHost}`);
        if (!proxyUri && !proxyUrl.port && config.proxy.port) {
            const port = Number(config.proxy.port);
            if (!Number.isSafeInteger(port) || port < 1 || port > 65535) {
                throw invalidProxy();
            }
            proxyUrl.port = String(port);
        }
    } catch {
        throw invalidProxy();
    }
    if (proxyUrl.protocol !== 'http:' && proxyUrl.protocol !== 'https:') {
        throw unsupportedProxy();
    }

    const token = !proxyUrl.username && !proxyUrl.password && config.proxy.auth ? `Basic ${config.proxy.auth}` : undefined;
    const key = JSON.stringify([proxyUrl.href, token]);
    let dispatcher = proxyDispatchers.get(key);
    if (!dispatcher) {
        try {
            dispatcher = new ProxyAgent({ uri: proxyUrl.href, ...(token && { token }) });
        } catch {
            throw invalidProxy();
        }
        proxyDispatchers.set(key, dispatcher);
    }
    return dispatcher;
};

export const postOAuthForm: OAuthRequest = async (url, form, options) => {
    const timeoutMs = Math.min(requestTimeoutMs, options?.timeoutMs ?? requestTimeoutMs);
    if (!Number.isFinite(timeoutMs) || timeoutMs <= 0) {
        throw new MailAuthError('Mail OAuth request has expired. Run the mail-auth command again.');
    }
    const controller = new AbortController();
    const signal = options?.signal ? AbortSignal.any([options.signal, controller.signal]) : controller.signal;
    const timeout = setTimeout(() => controller.abort(), timeoutMs);

    try {
        if (signal.aborted) {
            throw new MailAuthError('Mail OAuth request was cancelled. Retry the mail-auth command.');
        }

        const payload = Buffer.from(form.toString());
        const response = await undiciFetch(url, {
            method: 'POST',
            headers: { 'content-type': 'application/x-www-form-urlencoded', 'content-length': String(payload.length), accept: 'application/json' },
            // A one-shot body prevents Fetch's automatic replay of HTTP 421 responses.
            body: Readable.from([payload]),
            duplex: 'half',
            redirect: 'error',
            dispatcher: getProxyDispatcher(url),
            signal,
        });

        let body: unknown;
        try {
            body = await response.json();
        } catch {
            if (signal.aborted) {
                throw new Error('aborted');
            }
            throw new MailAuthError('Microsoft OAuth returned an invalid response. Check network/proxy settings and retry the mail-auth command.');
        }
        if (signal.aborted) {
            throw new Error('aborted');
        }
        return { status: response.status, body };
    } catch (error) {
        if (error instanceof MailAuthError) {
            throw error;
        }
        if (options?.signal?.aborted) {
            throw new MailAuthError('Mail OAuth request was cancelled. Retry the mail-auth command.');
        }
        if (controller.signal.aborted) {
            throw new MailAuthError('Microsoft OAuth request timed out. Check network/proxy settings and retry the mail-auth command.');
        }
        throw new MailAuthError('Microsoft OAuth request failed. Check network/proxy settings; use PROXY_STRATEGY=all if OAuth requires a proxy.');
    } finally {
        clearTimeout(timeout);
    }
};
