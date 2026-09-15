import type { Context } from 'hono';
import type { MailboxObject } from 'imapflow';
import PostalMime from 'postal-mime';

import { config } from '@/config';
import type { Route } from '@/types';
import cache from '@/utils/cache';
import { getImapAuth, parseMailAccount } from '@/utils/mail-auth';
import { withImapMailbox } from '@/utils/mail-auth/imap';
import { parseDate } from '@/utils/parse-date';

export const route: Route = {
    path: '/imap/:email/:folder{.+}?',
    categories: ['other'],
    example: '/mail/imap/rss@rsshub.app',
    parameters: {
        email: 'Email account',
        folder: 'Inbox name, `INBOX` by default',
    },
    description:
        "Supports IMAP with password or Outlook OAuth2 authentication. Email password and other settings refer to [Route-specific Configurations](https://docs.rsshub.app/deploy/config#route-specific-configurations); OAuth setup is documented in this fork's `docs/mail-oauth.zh.md`.",
    name: 'Inbox',
    maintainers: ['kt286'],
    handler,
};

async function handler(ctx: Context) {
    const { email, folder = 'INBOX' } = ctx.req.param();
    const { limit = 10 } = ctx.req.query();
    const account = parseMailAccount(email, config.email.config[email.replaceAll(/[.@]/g, '_')]);
    const auth = await getImapAuth(account, config.email.oauthStateDir);

    return withImapMailbox(account, auth, folder, async (client) => {
        const mails: any[] = [];
        try {
            const messages = client.fetch(`${Math.max((client.mailbox as MailboxObject).exists - Number(limit) + 1, 1)}:*`, { envelope: true, source: true, uid: true });
            for await (const message of messages) {
                mails.push(message);
            }
        } catch {
            throw new Error('Could not read messages from the IMAP mailbox. Check the connection and mailbox access, then retry.');
        }

        const items = await Promise.all(
            mails.map((item) =>
                cache.tryGet(`mail:${email}:${item.envelope.messageId}`, async () => {
                    const parsed = await PostalMime.parse(item.source);

                    let description = parsed.html || parsed.text?.replaceAll('\n', '<br>');
                    if (parsed.attachments.length) {
                        description += `<h3>Attachments (${parsed.attachments.length})</h3>`;
                        for (const attachment of parsed.attachments) {
                            description += `<p>${attachment.filename}</p>`;
                        }
                    }

                    return {
                        title: item.envelope.subject,
                        description,
                        pubDate: parseDate(item.envelope.date),
                        author: parsed.from!.name || parsed.from!.address,
                        guid: `mail:${email}:${item.envelope.messageId}`,
                    };
                })
            )
        );

        return {
            title: `${email}'s Inbox${folder === 'INBOX' ? '' : ` - ${folder}`}`,
            link: `https://${email.split('@', 2)[1]}`,
            item: items,
            allowEmpty: true,
        };
    });
}
