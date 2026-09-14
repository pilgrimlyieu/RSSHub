import ConfigNotFoundError from '@/errors/types/config-not-found';

interface MailAccountBase {
    email: string;
    username: string;
    host: string;
    port: number;
}

export interface PasswordMailAccount extends MailAccountBase {
    auth: 'password';
    password: string;
}

export interface OutlookMailAccount extends MailAccountBase {
    auth: 'oauth2';
    provider: 'outlook';
    clientId: string;
    tenant: string;
}

export type MailAccount = PasswordMailAccount | OutlookMailAccount;

export type ImapAuth = { user: string; pass: string } | { user: string; accessToken: string };

const guidPattern = /^[\da-f]{8}-[\da-f]{4}-[\da-f]{4}-[\da-f]{4}-[\da-f]{12}$/i;
const outlookTenants = new Set(['consumers', 'common', 'organizations']);

export const parseMailAccount = (email: string, rawConfig?: string): MailAccount => {
    const settings = new URLSearchParams(rawConfig);
    const auth = settings.get('auth') ?? 'password';
    const username = settings.get('username') ?? email;
    const port = Number(settings.get('port') ?? 993);
    const configKey = `EMAIL_CONFIG_${email.replaceAll(/[.@]/g, '_')}`;

    if (!username || !Number.isSafeInteger(port) || port < 1 || port > 65535) {
        throw new ConfigNotFoundError(`${configKey} must specify a username and a valid IMAP port (1-65535).`);
    }

    if (auth === 'password') {
        const host = settings.get('host');
        const password = settings.get('password');
        if (!host || !password) {
            throw new ConfigNotFoundError(`Configure host and password in ${configKey}, or select auth=oauth2 with a supported provider.`);
        }
        return { email, username, host, port, auth, password };
    }

    if (auth !== 'oauth2' || settings.get('provider') !== 'outlook') {
        throw new ConfigNotFoundError(`${configKey} must use auth=password or auth=oauth2&provider=outlook.`);
    }

    const clientId = settings.get('clientId') ?? '';
    const tenant = settings.get('tenant') ?? 'consumers';
    const host = settings.get('host') ?? 'outlook.office365.com';
    if (!guidPattern.test(clientId)) {
        throw new ConfigNotFoundError(`Set clientId in ${configKey} to your Microsoft application's client ID.`);
    }
    if (!outlookTenants.has(tenant) && !guidPattern.test(tenant)) {
        throw new ConfigNotFoundError(`Set tenant in ${configKey} to consumers for personal Outlook, organizations, common, or a directory tenant ID.`);
    }
    if (!host) {
        throw new ConfigNotFoundError(`Set a nonempty IMAP host in ${configKey}.`);
    }

    return { email, username, host, port, auth, provider: 'outlook', clientId, tenant };
};
