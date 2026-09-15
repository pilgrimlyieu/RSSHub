import ConfigNotFoundError from '@/errors/types/config-not-found';
import { isWorker } from '@/utils/is-worker';

import type { ImapAuth, MailAccount, OutlookMailAccount } from './account';
import { MailAuthError } from './errors';
import { verifyImapAuth } from './imap';
import { authorizeOutlook, type DeviceCodePrompt, refreshOutlook } from './outlook';
import { createState, prepareStateDirectory, readState, statePath, withStateLock, writeState } from './store';

export type { ImapAuth, MailAccount, OutlookMailAccount } from './account';
export { parseMailAccount } from './account';
export type { DeviceCodePrompt } from './outlook';

const checkRuntime = (account: OutlookMailAccount, stateDirectory?: string) => {
    if (isWorker) {
        throw new ConfigNotFoundError('Mail OAuth state requires Node.js with a persistent filesystem; use a Node or Docker RSSHub deployment.');
    }
    statePath(account, stateDirectory);
};

const missingAuthorization = (account: MailAccount) => new MailAuthError(`Mail OAuth authorization is missing. Run node dist/mail-auth.mjs login ${account.email} using the same configuration and state directory as RSSHub.`);

const pendingRefreshes = new Map<string, Promise<ImapAuth>>();

export const getImapAuth = async (account: MailAccount, stateDirectory?: string, options?: { forceRefresh?: boolean }): Promise<ImapAuth> => {
    if (account.auth === 'password') {
        return { user: account.username, pass: account.password };
    }
    checkRuntime(account, stateDirectory);
    const state = await readState(account, stateDirectory);
    if (!state) {
        throw missingAuthorization(account);
    }
    if (!options?.forceRefresh && state.expiresAt > Date.now() + 60000) {
        return { user: account.username, accessToken: state.accessToken };
    }

    const key = statePath(account, stateDirectory);
    let pending = pendingRefreshes.get(key);
    if (!pending) {
        pending = withStateLock(account, stateDirectory, async (lease) => {
            const latest = await readState(account, stateDirectory);
            if (!latest) {
                throw missingAuthorization(account);
            }
            if (!options?.forceRefresh && latest.expiresAt > Date.now() + 60000) {
                return { user: account.username, accessToken: latest.accessToken };
            }
            const refreshed = await refreshOutlook(account, latest.refreshToken, lease.signal);
            const next = createState(account, { ...refreshed, refreshToken: refreshed.refreshToken ?? latest.refreshToken });
            await writeState(account, stateDirectory, next, lease);
            return { user: account.username, accessToken: next.accessToken };
        });
        pendingRefreshes.set(key, pending);
    }
    try {
        return await pending;
    } finally {
        if (pendingRefreshes.get(key) === pending) {
            pendingRefreshes.delete(key);
        }
    }
};

export const authorizeMailAccount = async (account: MailAccount, stateDirectory: string | undefined, onDeviceCode: (prompt: DeviceCodePrompt) => void): Promise<{ messages: number }> => {
    if (account.auth !== 'oauth2') {
        throw new ConfigNotFoundError('This account uses password authentication. Configure auth=oauth2&provider=outlook and your clientId before running login.');
    }
    checkRuntime(account, stateDirectory);
    await prepareStateDirectory(account, stateDirectory);
    const tokens = await authorizeOutlook(account, onDeviceCode);
    // Prove that the new credentials can access the configured mailbox before replacing working state.
    const mailbox = await verifyImapAuth(account, { user: account.username, accessToken: tokens.accessToken });
    await withStateLock(account, stateDirectory, (lease) => writeState(account, stateDirectory, createState(account, tokens), lease));
    return mailbox;
};
