import type { DeviceCodePrompt } from '@/utils/mail-auth';

const usage = `Usage: node dist/mail-auth.mjs <login|check> <email>

  login <email>  Authorize OAuth access, verify read-only IMAP access, and save credentials.
  check <email>  Refresh OAuth credentials and verify read-only IMAP access.
  --help         Show this help without loading configuration or opening connections.

Configure EMAIL_CONFIG_<email> and EMAIL_OAUTH_STATE_DIR before using OAuth.
For password accounts, check verifies the configured IMAP credentials.
`;

function showDeviceCode({ verificationUri, userCode, expiresIn }: DeviceCodePrompt) {
    process.stdout.write(`Open this URL in your own browser: ${verificationUri}\nEnter this code: ${userCode}\nWaiting for authorization (expires in ${expiresIn} seconds).\n`);
}

async function main() {
    const args = process.argv.slice(2);
    if (args.length === 0 || (args.length === 1 && (args[0] === '--help' || args[0] === '-h'))) {
        process.stdout.write(usage);
        return;
    }

    const [command, email] = args;
    if (args.length !== 2 || (command !== 'login' && command !== 'check') || !email || email.startsWith('-')) {
        throw new Error('Usage: node dist/mail-auth.mjs <login|check> <email>. Use --help for details.');
    }

    const { config } = await import('@/config');
    const { authorizeMailAccount, getImapAuth, parseMailAccount } = await import('@/utils/mail-auth');
    const account = parseMailAccount(email, config.email.config[email.replaceAll(/[.@]/g, '_')]);
    const stateDir = config.email.oauthStateDir;

    if (command === 'login') {
        if (account.auth !== 'oauth2') {
            throw new Error('The login command requires auth=oauth2. Use check to verify a password account.');
        }
        const { messages } = await authorizeMailAccount(account, stateDir, showDeviceCode);
        process.stdout.write(`Authorization saved. Read-only INBOX check succeeded (${messages} messages).\n`);
        return;
    }

    const { verifyImapAuth } = await import('@/utils/mail-auth/imap');
    const auth = await getImapAuth(account, stateDir, { forceRefresh: true });
    const { messages } = await verifyImapAuth(account, auth);
    process.stdout.write(`${account.auth === 'oauth2' ? 'OAuth refresh and read-only INBOX check' : 'Read-only INBOX check'} succeeded (${messages} messages).\n`);
}

try {
    await main();
} catch (error) {
    process.stderr.write(`${error instanceof Error ? error.message : 'Mail authentication failed. Check the account configuration and retry.'}\n`);
    process.exitCode = 1;
}
