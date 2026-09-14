import { ImapFlow, type MailboxLockObject } from 'imapflow';

import { config } from '@/config';

import type { ImapAuth, MailAccount } from './account';

type ImapStage = 'connect' | 'mailbox' | 'connection';

const imapError = (account: MailAccount, stage: ImapStage, error: unknown): Error => {
    if (typeof error === 'object' && error !== null && 'authenticationFailed' in error && error.authenticationFailed) {
        return new Error(
            account.auth === 'oauth2'
                ? 'IMAP authentication failed. Check that you authorized the configured email account and enabled IMAP in the mailbox settings, then run mail-auth login again.'
                : 'IMAP authentication failed. Check the configured username, password and IMAP access in the mailbox settings.'
        );
    }

    return new Error(
        stage === 'mailbox' ? 'Could not open the IMAP mailbox. Check that the folder exists and the account has IMAP access.' : 'The IMAP connection failed. Check the configured host, port, proxy and network access, then retry.'
    );
};

const closeImapConnection = (client: ImapFlow): void => {
    try {
        client.close();
    } catch {
        // Cleanup failures must not replace the result or the original error.
    }
};

const cleanupImapMailbox = async (client: ImapFlow, connected: boolean, lock?: MailboxLockObject): Promise<void> => {
    try {
        lock?.release();
    } catch {
        closeImapConnection(client);
        return;
    }

    if (!connected) {
        closeImapConnection(client);
        return;
    }

    try {
        await client.logout();
    } catch {
        closeImapConnection(client);
    }
};

export const withImapMailbox = async <T>(account: MailAccount, auth: ImapAuth, folder: string, action: (client: ImapFlow) => Promise<T>): Promise<T> => {
    const client = new ImapFlow({
        host: account.host,
        port: account.port,
        secure: true,
        auth,
        proxy: config.proxyUri,
        logger: false,
    });
    let connectionError: Error | undefined;
    // Keep the listener after cleanup: late socket errors must remain handled.
    client.on('error', (error: unknown) => {
        connectionError ??= imapError(account, 'connection', error);
    });

    let connected = false;
    let lock: MailboxLockObject | undefined;
    try {
        try {
            await client.connect();
            connected = true;
        } catch (error) {
            throw imapError(account, 'connect', error);
        }
        if (connectionError) {
            throw connectionError;
        }

        try {
            lock = await client.getMailboxLock(folder, { readOnly: true });
        } catch (error) {
            throw imapError(account, 'mailbox', error);
        }
        if (connectionError) {
            throw connectionError;
        }

        const result = await action(client);
        if (connectionError) {
            throw connectionError;
        }
        return result;
    } finally {
        await cleanupImapMailbox(client, connected, lock);
    }
};

export const verifyImapAuth = (account: MailAccount, auth: ImapAuth): Promise<{ messages: number }> =>
    withImapMailbox(account, auth, 'INBOX', (client) => {
        if (!client.mailbox) {
            throw new Error('Could not open the IMAP inbox. Check that IMAP is enabled for the configured email account.');
        }
        return Promise.resolve({ messages: client.mailbox.exists });
    });
