import type { EventEmitter } from 'node:events';

import type { Context } from 'hono';
import type { FetchMessageObject, ImapFlowOptions } from 'imapflow';
import PostalMime from 'postal-mime';
import { afterEach, beforeEach, describe, expect, it, vi } from 'vitest';

import type { Data } from '../lib/types';
import { type MailAccount, parseMailAccount } from '../lib/utils/mail-auth/account';
import { verifyImapAuth, withImapMailbox } from '../lib/utils/mail-auth/imap';

const mocks = vi.hoisted(() => ({
    config: {
        email: { config: {} as Record<string, string>, oauthStateDir: '/private/mail-auth' },
        proxyUri: 'http://proxy.example:3128',
    },
    clients: [] as EventEmitter[],
    options: [] as ImapFlowOptions[],
    exists: 25,
    connect: vi.fn(),
    getMailboxLock: vi.fn(),
    release: vi.fn(),
    fetch: vi.fn(),
    logout: vi.fn(),
    close: vi.fn(),
    getImapAuth: vi.fn(),
    tryGet: vi.fn(),
}));

vi.mock('../lib/config', () => ({ config: mocks.config }));
vi.mock('../lib/utils/cache', () => ({ default: { tryGet: mocks.tryGet } }));
vi.mock('../lib/utils/mail-auth', async () => ({
    parseMailAccount: (await import('../lib/utils/mail-auth/account')).parseMailAccount,
    getImapAuth: mocks.getImapAuth,
}));
vi.mock('imapflow', async () => {
    const { EventEmitter } = await import('node:events');
    return {
        // oxlint-disable-next-line unicorn/prefer-event-target -- Match ImapFlow's Node error event semantics.
        ImapFlow: class extends EventEmitter {
            mailbox = { exists: mocks.exists };

            constructor(options: ImapFlowOptions) {
                super();
                mocks.clients.push(this);
                mocks.options.push(options);
            }

            connect() {
                return mocks.connect(this);
            }

            getMailboxLock(...args: unknown[]) {
                return mocks.getMailboxLock(...args);
            }

            fetch(...args: unknown[]) {
                return mocks.fetch(...args);
            }

            logout() {
                return mocks.logout(this);
            }

            close() {
                return mocks.close(this);
            }
        },
    };
});

const email = 'alice@outlook.com';
const passwordConfig = 'username=reader%40example.com&password=test-only-password&host=imap.example.com&port=9993';
const oauthConfig = 'auth=oauth2&provider=outlook&clientId=00000000-0000-4000-8000-000000000000';
const oauthAccount = parseMailAccount(email, oauthConfig);
const oauthAuth = { user: email, accessToken: 'test-only-token' };
const upstreamError = Object.assign(new Error('Upstream exposed test-only-token'), { responseText: 'test-only-token', response: { credentials: 'test-only-token' } });

const mail = {
    seq: 25,
    uid: 100,
    envelope: { subject: '  Newsletter subject  ', messageId: '<news@example.com>', date: new Date('2026-09-14T01:00:00Z') },
    source: Buffer.from(
        [
            'From: Newsletter <news@example.com>',
            'To: alice@outlook.com',
            'MIME-Version: 1.0',
            'Content-Type: multipart/mixed; boundary=example',
            '',
            '--example',
            'Content-Type: text/html; charset=utf-8',
            '',
            '<p>Newsletter body</p>',
            '--example',
            'Content-Type: text/plain; name=note.txt',
            'Content-Disposition: attachment; filename=note.txt',
            '',
            'Attachment text',
            '--example--',
            '',
        ].join('\r\n')
    ),
};

async function* fetchMessages(messages: FetchMessageObject[], error?: Error) {
    yield* messages;
    if (error) {
        throw error;
    }
}

const invokeRoute = async (folder?: string, query: Record<string, string> = {}) => {
    const { route } = await import('../lib/routes/mail/imap');
    return route.handler({ req: { param: () => ({ email, ...(folder && { folder }) }), query: () => query } } as unknown as Context) as Promise<Data>;
};

beforeEach(() => {
    vi.resetAllMocks();
    mocks.clients.length = 0;
    mocks.options.length = 0;
    mocks.exists = 25;
    mocks.config.email.config = { alice_outlook_com: passwordConfig };
    mocks.connect.mockResolvedValue(undefined);
    mocks.getMailboxLock.mockResolvedValue({ release: mocks.release });
    mocks.logout.mockResolvedValue(undefined);
    mocks.fetch.mockImplementation(() => fetchMessages([mail]));
    mocks.tryGet.mockImplementation((_key: string, callback: () => Promise<unknown>) => callback());
    mocks.getImapAuth.mockImplementation((account: MailAccount) => Promise.resolve(account.auth === 'password' ? { user: account.username, pass: account.password } : oauthAuth));
});

afterEach(() => {
    vi.restoreAllMocks();
});

describe('IMAP mail route', () => {
    it('preserves password authentication, existing RSS content and the default folder and limit', async () => {
        const result = await invokeRoute();

        expect(mocks.options).toEqual([{ host: 'imap.example.com', port: 9993, secure: true, auth: { user: 'reader@example.com', pass: 'test-only-password' }, proxy: mocks.config.proxyUri, logger: false }]);
        expect(mocks.getImapAuth).toHaveBeenCalledExactlyOnceWith(parseMailAccount(email, passwordConfig), '/private/mail-auth');
        expect(mocks.getMailboxLock).toHaveBeenCalledExactlyOnceWith('INBOX', { readOnly: true });
        expect(mocks.fetch).toHaveBeenCalledExactlyOnceWith('16:*', { envelope: true, source: true, uid: true });
        expect(mocks.tryGet).toHaveBeenCalledWith('mail:alice@outlook.com:<news@example.com>', expect.any(Function));
        expect(result).toEqual({
            title: "alice@outlook.com's Inbox",
            link: 'https://outlook.com',
            item: [
                {
                    title: '  Newsletter subject  ',
                    description: '<p>Newsletter body</p>\n<h3>Attachments (1)</h3><p>note.txt</p>',
                    pubDate: new Date('2026-09-14T01:00:00Z'),
                    author: 'Newsletter',
                    guid: 'mail:alice@outlook.com:<news@example.com>',
                },
            ],
            allowEmpty: true,
        });
        expect(mocks.release).toHaveBeenCalledOnce();
        expect(mocks.logout).toHaveBeenCalledOnce();
    });

    it('uses OAuth accessToken authentication and keeps custom folders and the common limit parameter', async () => {
        mocks.config.email.config.alice_outlook_com = oauthConfig;

        const result = await invokeRoute('Newsletters/Weekly', { limit: '2' });

        expect(mocks.options[0]).toMatchObject({ host: 'outlook.office365.com', port: 993, auth: oauthAuth, logger: false });
        expect(mocks.options[0].auth).not.toHaveProperty('pass');
        expect(mocks.getImapAuth).toHaveBeenCalledExactlyOnceWith(oauthAccount, '/private/mail-auth');
        expect(mocks.getMailboxLock).toHaveBeenCalledExactlyOnceWith('Newsletters/Weekly', { readOnly: true });
        expect(mocks.fetch).toHaveBeenCalledExactlyOnceWith('24:*', { envelope: true, source: true, uid: true });
        expect(result.title).toBe("alice@outlook.com's Inbox - Newsletters/Weekly");
    });

    it('reuses cached entries without parsing their source again', async () => {
        const cached = { title: 'Cached mail', description: '<p>Cached body</p>', guid: 'existing-guid' };
        mocks.tryGet.mockResolvedValue(cached);
        const parse = vi.spyOn(PostalMime, 'parse');

        expect((await invokeRoute()).item).toEqual([cached]);
        expect(parse).not.toHaveBeenCalled();
        expect(mocks.release).toHaveBeenCalledOnce();
        expect(mocks.logout).toHaveBeenCalledOnce();
    });

    it('does not open a connection if obtaining authentication fails', async () => {
        const failure = new Error('Run mail-auth login first.');
        mocks.getImapAuth.mockRejectedValue(failure);

        await expect(invokeRoute()).rejects.toBe(failure);
        expect(mocks.clients).toHaveLength(0);
    });

    it('sanitizes failed fetches and releases the connection instead of returning partial results', async () => {
        mocks.fetch.mockImplementation(() => fetchMessages([mail], upstreamError));

        const result = invokeRoute();

        await expect(result).rejects.toEqual(new Error('Could not read messages from the IMAP mailbox. Check the connection and mailbox access, then retry.'));
        await expect(result).rejects.not.toHaveProperty('cause');
        expect(mocks.tryGet).not.toHaveBeenCalled();
        expect(mocks.release).toHaveBeenCalledOnce();
        expect(mocks.logout).toHaveBeenCalledOnce();
    });

    it('preserves parser failures and still closes the connection when logout also fails', async () => {
        const parseError = new Error('The email body could not be parsed.');
        vi.spyOn(PostalMime, 'parse').mockRejectedValue(parseError);
        mocks.logout.mockRejectedValue(upstreamError);

        await expect(invokeRoute()).rejects.toBe(parseError);

        expect(mocks.release).toHaveBeenCalledOnce();
        expect(mocks.logout).toHaveBeenCalledOnce();
        expect(mocks.close).toHaveBeenCalledOnce();
    });
});

describe('IMAP connection lifecycle', () => {
    it('verifies authorization with a read-only INBOX without fetching message contents', async () => {
        await expect(verifyImapAuth(oauthAccount, oauthAuth)).resolves.toEqual({ messages: 25 });

        expect(mocks.options[0]).toMatchObject({ auth: oauthAuth, secure: true, logger: false });
        expect(mocks.getMailboxLock).toHaveBeenCalledExactlyOnceWith('INBOX', { readOnly: true });
        expect(mocks.fetch).not.toHaveBeenCalled();
        expect(mocks.release).toHaveBeenCalledOnce();
        expect(mocks.logout).toHaveBeenCalledOnce();
    });

    it('turns authentication failures into actionable errors without preserving upstream credentials', async () => {
        mocks.connect.mockRejectedValue(Object.assign(new Error('test-only-token'), { authenticationFailed: true, responseText: 'test-only-token', oauthError: { token: 'test-only-token' } }));

        const result = verifyImapAuth(oauthAccount, oauthAuth);

        await expect(result).rejects.toEqual(new Error('IMAP authentication failed. Check that you authorized the configured email account and enabled IMAP in the mailbox settings, then run mail-auth login again.'));
        await expect(result).rejects.not.toHaveProperty('cause');
        await expect(result).rejects.not.toHaveProperty('oauthError');
        expect(mocks.close).toHaveBeenCalledOnce();
        expect(mocks.getMailboxLock).not.toHaveBeenCalled();
        expect(mocks.logout).not.toHaveBeenCalled();
    });

    it('sanitizes network failures even if force-closing the connection also throws', async () => {
        mocks.connect.mockRejectedValue(upstreamError);
        mocks.close.mockImplementation(() => {
            throw upstreamError;
        });

        const result = verifyImapAuth(oauthAccount, oauthAuth);

        await expect(result).rejects.toEqual(new Error('The IMAP connection failed. Check the configured host, port, proxy and network access, then retry.'));
        await expect(result).rejects.not.toHaveProperty('cause');
        expect(mocks.close).toHaveBeenCalledOnce();
        expect(mocks.logout).not.toHaveBeenCalled();
    });

    it('closes a connected session when opening the mailbox fails', async () => {
        mocks.getMailboxLock.mockRejectedValue(upstreamError);

        const result = verifyImapAuth(oauthAccount, oauthAuth);

        await expect(result).rejects.toEqual(new Error('Could not open the IMAP mailbox. Check that the folder exists and the account has IMAP access.'));
        await expect(result).rejects.not.toHaveProperty('cause');
        expect(mocks.release).not.toHaveBeenCalled();
        expect(mocks.logout).toHaveBeenCalledOnce();
    });

    it('handles error events emitted during connect before any mailbox action starts', async () => {
        mocks.connect.mockImplementation((client: EventEmitter) => {
            client.emit('error', upstreamError);
            return Promise.resolve();
        });

        await expect(verifyImapAuth(oauthAccount, oauthAuth)).rejects.toThrow('The IMAP connection failed.');

        expect(mocks.getMailboxLock).not.toHaveBeenCalled();
        expect(mocks.logout).toHaveBeenCalledOnce();
    });

    it('rejects results when the connection emits an asynchronous error during the action', async () => {
        await expect(
            withImapMailbox(oauthAccount, oauthAuth, 'INBOX', async (client) => {
                await Promise.resolve();
                client.emit('error', upstreamError);
                return 'Incomplete result';
            })
        ).rejects.toThrow('The IMAP connection failed.');

        expect(mocks.release).toHaveBeenCalledOnce();
        expect(mocks.logout).toHaveBeenCalledOnce();
    });

    it('force-closes after a failed logout and safely receives later error events', async () => {
        mocks.logout.mockRejectedValue(upstreamError);

        await expect(verifyImapAuth(oauthAccount, oauthAuth)).resolves.toEqual({ messages: 25 });

        expect(mocks.release).toHaveBeenCalledOnce();
        expect(mocks.close).toHaveBeenCalledOnce();
        expect(() => mocks.clients[0].emit('error', upstreamError)).not.toThrow();
    });

    it('still closes the connection when releasing the mailbox lock fails', async () => {
        mocks.release.mockImplementation(() => {
            throw upstreamError;
        });

        await expect(verifyImapAuth(oauthAccount, oauthAuth)).resolves.toEqual({ messages: 25 });

        expect(mocks.close).toHaveBeenCalledOnce();
        expect(mocks.logout).not.toHaveBeenCalled();
    });
});
