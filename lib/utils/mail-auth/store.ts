import { createHash, randomUUID } from 'node:crypto';
import { mkdir, open, readFile, rename, rm } from 'node:fs/promises';
import path from 'node:path';

import ConfigNotFoundError from '@/errors/types/config-not-found';

import type { OutlookMailAccount } from './account';
import { MailAuthError } from './errors';

interface OAuthIdentity {
    provider: 'outlook';
    username: string;
    clientId: string;
    tenant: string;
}

export interface OAuthState {
    version: 1;
    identity: OAuthIdentity;
    accessToken: string;
    refreshToken: string;
    expiresAt: number;
}

export interface StateLease {
    signal: AbortSignal;
    assertActive: () => void;
}

const cleanupTemporaryFile = async (file: string, handle?: Awaited<ReturnType<typeof open>>): Promise<void> => {
    try {
        await handle?.close();
    } catch {
        // Preserve the original storage error; the file may already be closed.
    }
    try {
        await rm(file, { force: true });
    } catch {
        // A leftover private temporary file must not replace the original error.
    }
};

const identityOf = (account: OutlookMailAccount): OAuthIdentity => ({
    provider: account.provider,
    username: account.username.toLowerCase(),
    clientId: account.clientId.toLowerCase(),
    tenant: account.tenant.toLowerCase(),
});

export const statePath = (account: OutlookMailAccount, directory?: string): string => {
    if (!directory) {
        throw new ConfigNotFoundError('Set EMAIL_OAUTH_STATE_DIR to a persistent, writable directory before using mail OAuth.');
    }
    const key = createHash('sha256')
        .update(JSON.stringify(identityOf(account)))
        .digest('hex');
    return path.join(path.resolve(directory), `${key}.json`);
};

export const createState = (account: OutlookMailAccount, tokens: { accessToken: string; refreshToken: string; expiresAt: number }): OAuthState => ({
    version: 1,
    identity: identityOf(account),
    accessToken: tokens.accessToken,
    refreshToken: tokens.refreshToken,
    expiresAt: tokens.expiresAt,
});

export const prepareStateDirectory = async (account: OutlookMailAccount, directory?: string): Promise<void> => {
    const file = statePath(account, directory);
    const probe = `${file}.${randomUUID()}.tmp`;
    let handle: Awaited<ReturnType<typeof open>> | undefined;
    try {
        await mkdir(path.dirname(file), { recursive: true, mode: 0o700 });
        handle = await open(probe, 'wx', 0o600);
    } catch {
        throw new MailAuthError('Mail OAuth requires a writable EMAIL_OAUTH_STATE_DIR. Check the Docker mount and directory permissions before authorizing.');
    } finally {
        await cleanupTemporaryFile(probe, handle);
    }
};

const isRecord = (value: unknown): value is Record<string, unknown> => typeof value === 'object' && value !== null && !Array.isArray(value);

const parseState = (account: OutlookMailAccount, raw: string): OAuthState => {
    const value: unknown = JSON.parse(raw);
    const identity = identityOf(account);
    if (
        !isRecord(value) ||
        value.version !== 1 ||
        !isRecord(value.identity) ||
        Object.entries(identity).some(([key, expected]) => (value.identity as Record<string, unknown>)[key] !== expected) ||
        typeof value.accessToken !== 'string' ||
        !value.accessToken ||
        typeof value.refreshToken !== 'string' ||
        !value.refreshToken ||
        typeof value.expiresAt !== 'number' ||
        !Number.isSafeInteger(value.expiresAt) ||
        value.expiresAt <= 0
    ) {
        throw new MailAuthError('Invalid OAuth state.');
    }
    return createState(account, { accessToken: value.accessToken, refreshToken: value.refreshToken, expiresAt: value.expiresAt });
};

export const readState = async (account: OutlookMailAccount, directory?: string): Promise<OAuthState | undefined> => {
    const file = statePath(account, directory);
    let raw: string;
    try {
        raw = await readFile(file, 'utf8');
    } catch (error) {
        if ((error as NodeJS.ErrnoException).code === 'ENOENT') {
            return;
        }
        throw new MailAuthError('Unable to read mail OAuth state. Check EMAIL_OAUTH_STATE_DIR and its filesystem permissions.');
    }
    try {
        return parseState(account, raw);
    } catch {
        throw new MailAuthError(`Mail OAuth state is damaged or does not match this account. Run node dist/mail-auth.mjs login ${account.email} to replace it.`);
    }
};

export const writeState = async (account: OutlookMailAccount, directory: string | undefined, state: OAuthState, lease: StateLease): Promise<void> => {
    const file = statePath(account, directory);
    const temporary = `${file}.${randomUUID()}.tmp`;
    let handle: Awaited<ReturnType<typeof open>> | undefined;
    try {
        lease.assertActive();
        handle = await open(temporary, 'wx', 0o600);
        await handle.writeFile(JSON.stringify(state), 'utf8');
        await handle.sync();
        await handle.close();
        handle = undefined;
        lease.assertActive();
        await rename(temporary, file);
    } catch {
        throw new MailAuthError('Unable to save mail OAuth state. Check that EMAIL_OAUTH_STATE_DIR is writable and retry authorization if necessary.');
    } finally {
        await cleanupTemporaryFile(temporary, handle);
    }
};

export const withStateLock = async <T>(account: OutlookMailAccount, directory: string | undefined, action: (lease: StateLease) => Promise<T>): Promise<T> => {
    const file = statePath(account, directory);
    try {
        await mkdir(path.dirname(file), { recursive: true, mode: 0o700 });
    } catch {
        throw new MailAuthError('Unable to create the mail OAuth state directory. Check EMAIL_OAUTH_STATE_DIR and its filesystem permissions.');
    }

    // Load filesystem locking only when OAuth is used; password routes need no writable state.
    const { lock } = await import('proper-lockfile');
    const controller = new AbortController();
    let compromised = false;
    let release: () => Promise<void>;
    try {
        release = await lock(file, {
            realpath: false,
            stale: 30000,
            update: 5000,
            retries: { retries: 15, factor: 1, minTimeout: 1000, maxTimeout: 1000 },
            onCompromised: () => {
                compromised = true;
                controller.abort();
            },
        });
    } catch {
        throw new MailAuthError('Unable to lock mail OAuth state. Check directory permissions or retry after another authorization or refresh completes.');
    }

    const assertActive = () => {
        if (compromised) {
            throw new MailAuthError('The mail OAuth state lock was lost. Retry after other processes finish; no further state will be written by this operation.');
        }
    };
    let result: T | undefined;
    let failure: unknown;
    let succeeded = false;
    try {
        assertActive();
        result = await action({ signal: controller.signal, assertActive });
        assertActive();
        succeeded = true;
    } catch (error) {
        failure = error;
    }

    // A compromised lock may already belong to another process; do not release that lock.
    if (!compromised) {
        try {
            await release();
        } catch {
            if (succeeded) {
                throw new MailAuthError('Unable to release the mail OAuth state lock. Retry after the lock expires.');
            }
        }
    }
    if (!succeeded) {
        throw failure;
    }
    return result as T;
};
