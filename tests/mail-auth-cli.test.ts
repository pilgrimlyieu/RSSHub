import { afterEach, beforeEach, describe, expect, it, vi } from 'vitest';

const mocks = vi.hoisted(() => ({
    configLoaded: vi.fn(),
    parseMailAccount: vi.fn(),
    authorizeMailAccount: vi.fn(),
    getImapAuth: vi.fn(),
    verifyImapAuth: vi.fn(),
}));

vi.mock('../lib/config', () => {
    mocks.configLoaded();
    return {
        config: {
            email: {
                config: { alice_outlook_com: 'auth=oauth2&provider=outlook&clientId=example' },
                oauthStateDir: '/tmp/mail-auth-cli-state',
            },
        },
    };
});
vi.mock('../lib/utils/mail-auth', () => ({
    parseMailAccount: mocks.parseMailAccount,
    authorizeMailAccount: mocks.authorizeMailAccount,
    getImapAuth: mocks.getImapAuth,
}));
vi.mock('../lib/utils/mail-auth/imap', () => ({ verifyImapAuth: mocks.verifyImapAuth }));
vi.mock('../lib/app', () => {
    throw new Error('The mail CLI must not load the web app.');
});
vi.mock('../lib/utils/request-rewriter', () => {
    throw new Error('The mail CLI must not load the global request rewriter.');
});

const originalArgv = process.argv;
const originalExitCode = process.exitCode;
const email = 'alice@outlook.com';
const stateDir = '/tmp/mail-auth-cli-state';
const account = { email, username: email, host: 'outlook.office365.com', port: 993, auth: 'oauth2', provider: 'outlook', clientId: 'example', tenant: 'consumers' };
const auth = { user: email, accessToken: 'private-access-token' };
let stdout: string;
let stderr: string;

beforeEach(() => {
    vi.resetModules();
    vi.resetAllMocks();
    stdout = '';
    stderr = '';
    process.exitCode = undefined;
    vi.spyOn(process.stdout, 'write').mockImplementation((chunk) => {
        stdout += String(chunk);
        return true;
    });
    vi.spyOn(process.stderr, 'write').mockImplementation((chunk) => {
        stderr += String(chunk);
        return true;
    });
    mocks.parseMailAccount.mockReturnValue(account);
    mocks.authorizeMailAccount.mockResolvedValue({ messages: 3 });
    mocks.getImapAuth.mockResolvedValue(auth);
    mocks.verifyImapAuth.mockResolvedValue({ messages: 3 });
});

afterEach(() => {
    process.argv = originalArgv;
    process.exitCode = originalExitCode;
    vi.restoreAllMocks();
});

async function runCli(...args: string[]) {
    process.argv = [process.execPath, 'mail-auth.mjs', ...args];
    await import('../lib/cli/mail-auth');
}

describe('mail authorization CLI', () => {
    it('shows help without loading configuration or starting authentication', async () => {
        await runCli('--help');

        expect(stdout).toContain('Usage: node dist/mail-auth.mjs <login|check> <email>');
        expect(stderr).toBe('');
        expect(process.exitCode).toBeUndefined();
        expect(mocks.configLoaded).not.toHaveBeenCalled();
        expect(mocks.parseMailAccount).not.toHaveBeenCalled();
        expect(mocks.authorizeMailAccount).not.toHaveBeenCalled();
        expect(mocks.getImapAuth).not.toHaveBeenCalled();
    });

    it.each([['login'], ['unknown', email], ['check', email, 'extra']])('rejects malformed arguments before loading configuration: %j', async (...args) => {
        await runCli(...args);

        expect(stdout).toBe('');
        expect(stderr).toContain('Usage:');
        expect(process.exitCode).toBe(1);
        expect(mocks.configLoaded).not.toHaveBeenCalled();
        expect(mocks.parseMailAccount).not.toHaveBeenCalled();
    });

    it('shows the device prompt and waits for verified authorization to be saved', async () => {
        mocks.authorizeMailAccount.mockImplementation((_account, _stateDir, onDeviceCode) => {
            onDeviceCode({ verificationUri: 'https://microsoft.com/devicelogin', userCode: 'ABCD-EFGH', expiresIn: 900 });
            return Promise.resolve({ messages: 3, accessToken: 'private-access-token', refreshToken: 'private-refresh-token' });
        });

        await runCli('login', email);

        expect(mocks.parseMailAccount).toHaveBeenCalledExactlyOnceWith(email, 'auth=oauth2&provider=outlook&clientId=example');
        expect(mocks.authorizeMailAccount).toHaveBeenCalledExactlyOnceWith(account, stateDir, expect.any(Function));
        expect(stdout).toContain('https://microsoft.com/devicelogin');
        expect(stdout).toContain('ABCD-EFGH');
        expect(stdout).toContain('Authorization saved. Read-only INBOX check succeeded (3 messages).');
        expect(stdout).not.toContain('private-');
        expect(stderr).toBe('');
        expect(mocks.getImapAuth).not.toHaveBeenCalled();
    });

    it('forces a refresh before checking IMAP and never prints the access token', async () => {
        await runCli('check', email);

        expect(mocks.getImapAuth).toHaveBeenCalledExactlyOnceWith(account, stateDir, { forceRefresh: true });
        expect(mocks.verifyImapAuth).toHaveBeenCalledExactlyOnceWith(account, auth);
        expect(stdout).toBe('OAuth refresh and read-only INBOX check succeeded (3 messages).\n');
        expect(stderr).toBe('');
        expect(mocks.authorizeMailAccount).not.toHaveBeenCalled();
    });

    it('supports password checks without claiming an OAuth refresh occurred', async () => {
        const passwordAccount = { ...account, auth: 'password', password: 'private-password' };
        mocks.parseMailAccount.mockReturnValue(passwordAccount);
        mocks.getImapAuth.mockResolvedValue({ user: email, pass: 'private-password' });

        await runCli('check', email);

        expect(mocks.verifyImapAuth).toHaveBeenCalledExactlyOnceWith(passwordAccount, { user: email, pass: 'private-password' });
        expect(stdout).toBe('Read-only INBOX check succeeded (3 messages).\n');
        expect(stderr).toBe('');
    });

    it('requires an OAuth account for login', async () => {
        mocks.parseMailAccount.mockReturnValue({ ...account, auth: 'password', password: 'private-password' });

        await runCli('login', email);

        expect(stderr).toContain('The login command requires auth=oauth2.');
        expect(process.exitCode).toBe(1);
        expect(mocks.authorizeMailAccount).not.toHaveBeenCalled();
    });

    it('prints only the safe error message and exits unsuccessfully', async () => {
        const failure = new Error('Authorization must be renewed. Run login again.', { cause: { refreshToken: 'private-refresh-token' } });
        mocks.getImapAuth.mockRejectedValue(failure);

        await runCli('check', email);

        expect(stdout).toBe('');
        expect(stderr).toBe('Authorization must be renewed. Run login again.\n');
        expect(process.exitCode).toBe(1);
        expect(mocks.verifyImapAuth).not.toHaveBeenCalled();
    });
});
