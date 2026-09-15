import { chmod, mkdir, mkdtemp, readdir, readFile, rm, writeFile } from 'node:fs/promises';
import os from 'node:os';
import path from 'node:path';

import { afterEach, beforeEach, describe, expect, it, vi } from 'vitest';

import type * as mailAuthModule from '../lib/utils/mail-auth';
import { type OutlookMailAccount, parseMailAccount } from '../lib/utils/mail-auth/account';
import type { OAuthTokenSet } from '../lib/utils/mail-auth/outlook';
import { createState, readState, statePath, withStateLock, writeState } from '../lib/utils/mail-auth/store';

const mocks = vi.hoisted(() => ({ authorizeOutlook: vi.fn(), refreshOutlook: vi.fn(), verifyImapAuth: vi.fn() }));

vi.mock('../lib/utils/mail-auth/outlook', () => ({ authorizeOutlook: mocks.authorizeOutlook, refreshOutlook: mocks.refreshOutlook }));
vi.mock('../lib/utils/mail-auth/imap', () => ({ verifyImapAuth: mocks.verifyImapAuth }));

const email = 'alice@outlook.com';
const clientId = '11111111-1111-4111-8111-111111111111';
const oauthConfig = `auth=oauth2&provider=outlook&clientId=${clientId}`;
const account: OutlookMailAccount = { email, username: email, host: 'outlook.office365.com', port: 993, auth: 'oauth2', provider: 'outlook', clientId, tenant: 'consumers' };
let root: string;
let directory: string;
let now: number;
let mailAuth: typeof mailAuthModule;

beforeEach(async () => {
    vi.resetModules();
    vi.resetAllMocks();
    root = await mkdtemp(path.join(os.tmpdir(), 'rsshub-mail-auth-'));
    directory = path.join(root, 'state');
    now = Date.now();
    mocks.authorizeOutlook.mockResolvedValue({ accessToken: 'authorized-access', refreshToken: 'authorized-refresh', expiresAt: now + 3_600_000 });
    mocks.refreshOutlook.mockResolvedValue({ accessToken: 'refreshed-access', refreshToken: 'rotated-refresh', expiresAt: now + 3_600_000 });
    mocks.verifyImapAuth.mockResolvedValue({ messages: 4 });
    mailAuth = await import('../lib/utils/mail-auth');
});

afterEach(async () => {
    vi.restoreAllMocks();
    await rm(root, { recursive: true, force: true });
});

async function saveState(overrides: Partial<OAuthTokenSet> = {}, identity = account) {
    const state = createState(identity, { accessToken: 'original-access', refreshToken: 'original-refresh', expiresAt: now + 3_600_000, ...overrides });
    await withStateLock(identity, directory, (lease) => writeState(identity, directory, state, lease));
    return state;
}

describe('mail account configuration', () => {
    it('keeps the password defaults and decodes the existing query-string format', async () => {
        const password = parseMailAccount(email, 'host=imap.example.com&password=private%26password');

        expect(password).toEqual({ email, username: email, host: 'imap.example.com', port: 993, auth: 'password', password: 'private&password' });
        expect(await mailAuth.getImapAuth(password)).toEqual({ user: email, pass: 'private&password' });
        expect(await mailAuth.getImapAuth(password, '/not/a/writable/state/path', { forceRefresh: true })).toEqual({ user: email, pass: 'private&password' });
        expect(mocks.refreshOutlook).not.toHaveBeenCalled();
        await expect(readdir(directory)).rejects.toMatchObject({ code: 'ENOENT' });
    });

    it('preserves an explicitly configured IMAP username and port', () => {
        expect(parseMailAccount(email, 'host=imap.example.com&password=password&username=mailbox%40example.com&port=1993')).toMatchObject({ username: 'mailbox@example.com', port: 1993 });
    });

    it('defaults Outlook OAuth to the personal-account endpoint and TLS IMAP', () => {
        expect(parseMailAccount(email, oauthConfig)).toEqual(account);
    });

    it.each([
        ['unknown authentication', 'auth=magic&host=imap.example.com&password=secret'],
        ['unsupported provider', `auth=oauth2&provider=gmail&clientId=${clientId}`],
        ['missing provider', `auth=oauth2&clientId=${clientId}`],
        ['missing client ID', 'auth=oauth2&provider=outlook'],
        ['invalid client ID', 'auth=oauth2&provider=outlook&clientId=not-an-id'],
        ['tenant URL', `${oauthConfig}&tenant=https%3A%2F%2Fexample.com`],
        ['tenant path', `${oauthConfig}&tenant=consumers%2F..%2Fcommon`],
        ['empty username', `${oauthConfig}&username=`],
        ['empty host', `${oauthConfig}&host=`],
        ['invalid port', `${oauthConfig}&port=invalid`],
        ['zero port', `${oauthConfig}&port=0`],
        ['out-of-range port', `${oauthConfig}&port=65536`],
    ])('rejects %s without falling back to password authentication', (_name, raw) => {
        expect(() => parseMailAccount(email, raw)).toThrow();
        expect(mocks.authorizeOutlook).not.toHaveBeenCalled();
        expect(mocks.refreshOutlook).not.toHaveBeenCalled();
    });

    it('requires an explicit OAuth state directory and an existing authorization', async () => {
        await expect(mailAuth.getImapAuth(account)).rejects.toThrow('EMAIL_OAUTH_STATE_DIR');
        await expect(mailAuth.getImapAuth(account, directory)).rejects.toThrow(`login ${email}`);
        expect(mocks.refreshOutlook).not.toHaveBeenCalled();
        expect(mocks.authorizeOutlook).not.toHaveBeenCalled();
    });
});

describe('mail OAuth authorization', () => {
    it('verifies the configured mailbox before first saving credentials', async () => {
        mocks.verifyImapAuth.mockImplementation(async () => {
            expect(await readState(account, directory)).toBeUndefined();
            expect(await readdir(directory)).toEqual([]);
            return { messages: 4 };
        });
        const onDeviceCode = vi.fn();

        expect(await mailAuth.authorizeMailAccount(account, directory, onDeviceCode)).toEqual({ messages: 4 });

        expect(mocks.authorizeOutlook).toHaveBeenCalledExactlyOnceWith(account, onDeviceCode);
        expect(mocks.verifyImapAuth).toHaveBeenCalledExactlyOnceWith(account, { user: email, accessToken: 'authorized-access' });
        expect(await readState(account, directory)).toMatchObject({ accessToken: 'authorized-access', refreshToken: 'authorized-refresh', expiresAt: now + 3_600_000 });
    });

    it('keeps the existing file unchanged if the new grant fails IMAP verification', async () => {
        await saveState();
        const previous = await readFile(statePath(account, directory), 'utf8');
        const failure = new Error('The authorized account does not match the configured mailbox.');
        mocks.verifyImapAuth.mockRejectedValue(failure);

        await expect(mailAuth.authorizeMailAccount(account, directory, vi.fn())).rejects.toBe(failure);

        expect(await readFile(statePath(account, directory), 'utf8')).toBe(previous);
        expect(await readdir(directory)).toEqual([path.basename(statePath(account, directory))]);
    });

    it('can replace damaged state after the new authorization passes IMAP verification', async () => {
        await mkdir(directory);
        await writeFile(statePath(account, directory), 'damaged-private-state');

        await mailAuth.authorizeMailAccount(account, directory, vi.fn());

        expect(await readState(account, directory)).toMatchObject({ accessToken: 'authorized-access', refreshToken: 'authorized-refresh' });
    });

    it('does not hold the state lock while waiting for the user to authorize', async () => {
        const started = Promise.withResolvers<void>();
        const grant = Promise.withResolvers<OAuthTokenSet & { refreshToken: string }>();
        mocks.authorizeOutlook.mockImplementation(() => {
            started.resolve();
            return grant.promise;
        });
        const authorization = mailAuth.authorizeMailAccount(account, directory, vi.fn());
        await started.promise;

        const acquired = await withStateLock(account, directory, (lease) => {
            lease.assertActive();
            return Promise.resolve('lock is available');
        });
        grant.resolve({ accessToken: 'authorized-access', refreshToken: 'authorized-refresh', expiresAt: now + 3_600_000 });
        await authorization;

        expect(acquired).toBe('lock is available');
    });

    it('rejects password-account login without starting OAuth', async () => {
        const password = parseMailAccount(email, 'host=imap.example.com&password=password');

        await expect(mailAuth.authorizeMailAccount(password, directory, vi.fn())).rejects.toThrow('password authentication');

        expect(mocks.authorizeOutlook).not.toHaveBeenCalled();
    });

    it('fails before device authorization if the state path is a file', async () => {
        await writeFile(directory, 'existing file');

        await expect(mailAuth.authorizeMailAccount(account, directory, vi.fn())).rejects.toThrow('writable EMAIL_OAUTH_STATE_DIR');

        expect(mocks.authorizeOutlook).not.toHaveBeenCalled();
        expect(await readFile(directory, 'utf8')).toBe('existing file');
    });

    it.skipIf(process.platform === 'win32' || (process.getuid?.() === 0 && process.platform !== 'linux'))('fails before device authorization if the state directory is not writable', async () => {
        await mkdir(directory);
        await chmod(directory, 0o500);
        // Root bypasses POSIX mode bits; procfs still rejects creating an OAuth state directory.
        const unwritable = process.getuid?.() === 0 ? `/proc/self/${path.basename(root)}` : directory;
        try {
            await expect(mailAuth.authorizeMailAccount(account, unwritable, vi.fn())).rejects.toThrow('writable EMAIL_OAUTH_STATE_DIR');
            expect(mocks.authorizeOutlook).not.toHaveBeenCalled();
        } finally {
            await chmod(directory, 0o700);
        }
    });
});

describe('mail OAuth refresh', () => {
    it('uses absolute expiry, refreshes at the 60-second boundary, and does not extend expiry on reads', async () => {
        const original = await saveState({ expiresAt: now + 120000 });
        const currentTime = vi.spyOn(Date, 'now').mockReturnValue(now);

        expect(await mailAuth.getImapAuth(account, directory)).toEqual({ user: email, accessToken: original.accessToken });
        currentTime.mockReturnValue(now + 59000);
        expect(await mailAuth.getImapAuth(account, directory)).toEqual({ user: email, accessToken: original.accessToken });
        expect(await readState(account, directory)).toEqual(original);
        expect(mocks.refreshOutlook).not.toHaveBeenCalled();

        currentTime.mockReturnValue(now + 60000);
        expect(await mailAuth.getImapAuth(account, directory)).toEqual({ user: email, accessToken: 'refreshed-access' });
        expect(mocks.refreshOutlook).toHaveBeenCalledExactlyOnceWith(account, 'original-refresh', expect.any(AbortSignal));
        const refreshed = await readState(account, directory);
        currentTime.mockReturnValue(now + 90000);
        await mailAuth.getImapAuth(account, directory);
        expect(await readState(account, directory)).toEqual(refreshed);
        expect(refreshed?.expiresAt).toBe(now + 3_600_000);
        expect(mocks.refreshOutlook).toHaveBeenCalledOnce();
    });

    it.each([
        ['rotates a returned refresh token', 'replacement-refresh', 'replacement-refresh'],
        ['keeps the old refresh token when none is returned', undefined, 'original-refresh'],
    ])('%s', async (_name, refreshToken, expected) => {
        await saveState({ expiresAt: now - 1 });
        mocks.refreshOutlook.mockResolvedValue({ accessToken: 'refreshed-access', expiresAt: now + 3_600_000, ...(refreshToken && { refreshToken }) });

        expect(await mailAuth.getImapAuth(account, directory)).toEqual({ user: email, accessToken: 'refreshed-access' });

        expect(await readState(account, directory)).toMatchObject({ accessToken: 'refreshed-access', refreshToken: expected });
        await mailAuth.getImapAuth(account, directory, { forceRefresh: true });
        expect(mocks.refreshOutlook).toHaveBeenLastCalledWith(account, expected, expect.any(AbortSignal));
    });

    it('forces a real refresh even when the stored access token is still valid', async () => {
        await saveState();

        await mailAuth.getImapAuth(account, directory, { forceRefresh: true });

        expect(mocks.refreshOutlook).toHaveBeenCalledOnce();
        expect(await readState(account, directory)).toMatchObject({ accessToken: 'refreshed-access', refreshToken: 'rotated-refresh' });
    });

    it('shares one refresh between concurrent callers', async () => {
        await saveState({ expiresAt: now - 1 });
        const started = Promise.withResolvers<void>();
        const refresh = Promise.withResolvers<OAuthTokenSet>();
        mocks.refreshOutlook.mockImplementation(() => {
            started.resolve();
            return refresh.promise;
        });
        const requests = Promise.all(Array.from({ length: 8 }, () => mailAuth.getImapAuth(account, directory)));
        await started.promise;
        refresh.resolve({ accessToken: 'shared-access', refreshToken: 'shared-refresh', expiresAt: now + 3_600_000 });

        expect(await requests).toEqual(Array.from({ length: 8 }, () => ({ user: email, accessToken: 'shared-access' })));
        expect(mocks.refreshOutlook).toHaveBeenCalledOnce();
        expect(await readState(account, directory)).toMatchObject({ accessToken: 'shared-access', refreshToken: 'shared-refresh' });
    });

    it('preserves state after a shared refresh failure and allows the next call to retry', async () => {
        const original = await saveState({ expiresAt: now - 1 });
        const started = Promise.withResolvers<void>();
        const refresh = Promise.withResolvers<OAuthTokenSet>();
        mocks.refreshOutlook.mockImplementation(() => {
            started.resolve();
            return refresh.promise;
        });
        const requests = Promise.allSettled(Array.from({ length: 8 }, () => mailAuth.getImapAuth(account, directory)));
        await started.promise;
        refresh.reject(new Error('Refresh temporarily unavailable.'));

        expect((await requests).every((result) => result.status === 'rejected')).toBe(true);
        expect(mocks.refreshOutlook).toHaveBeenCalledOnce();
        expect(await readState(account, directory)).toEqual(original);

        mocks.refreshOutlook.mockResolvedValue({ accessToken: 'recovered-access', expiresAt: now + 3_600_000 });
        expect(await mailAuth.getImapAuth(account, directory)).toEqual({ user: email, accessToken: 'recovered-access' });
        expect(mocks.refreshOutlook).toHaveBeenCalledTimes(2);
    });

    it('rereads credentials after a module reload and after another authorization writes new state', async () => {
        await saveState();
        expect(await mailAuth.getImapAuth(account, directory)).toEqual({ user: email, accessToken: 'original-access' });
        vi.resetModules();
        const reloaded = await import('../lib/utils/mail-auth');
        expect(await reloaded.getImapAuth(account, directory)).toEqual({ user: email, accessToken: 'original-access' });

        await saveState({ accessToken: 'other-process-access', refreshToken: 'other-process-refresh' });

        expect(await mailAuth.getImapAuth(account, directory)).toEqual({ user: email, accessToken: 'other-process-access' });
        expect(await reloaded.getImapAuth(account, directory)).toEqual({ user: email, accessToken: 'other-process-access' });
        expect(mocks.refreshOutlook).not.toHaveBeenCalled();
    });

    it.each([
        ['client ID', { clientId: '22222222-2222-4222-8222-222222222222' }],
        ['IMAP username', { username: 'other@outlook.com' }],
        ['tenant', { tenant: 'common' }],
    ])('requires a separate authorization after changing %s', async (_name, change) => {
        const original = await saveState();
        const changed = { ...account, ...change };

        await expect(mailAuth.getImapAuth(changed, directory)).rejects.toThrow('authorization is missing');

        expect(await readState(account, directory)).toEqual(original);
        expect(mocks.refreshOutlook).not.toHaveBeenCalled();
    });
});
