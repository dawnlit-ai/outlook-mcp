// Bounce-backs: finding the non-delivery notices in a mailbox, measuring what
// was sent against them, and clearing them out. What to do about an address
// that bounced (stop mailing it, block a sender, fix a contact) is the host's
// own rule, so it isn't here.
import { z } from 'zod';
import { json, toolHandler } from '../results';
import type { ToolContext } from './context';

export function registerBounceTools({ server, bridge, readOnly, add }: ToolContext): void {
    const actions = readOnly ? (['list', 'report'] as const) : (['list', 'report', 'clean'] as const);
    add('outlook_bounces', 'read', () => server.registerTool('outlook_bounces', {
        title: 'Outlook bounce-backs',
        description:
            "The bounce-backs in an account: non-delivery reports, mailer-daemon and postmaster notices, and subjects only a mail system writes ('Undeliverable:', 'Delivery Status Notification (Failure)', …). Ordinary mail that merely mentions a delivery is never matched. " +
            "'list' finds them in the Inbox, with why each matched and the addresses it reports as failed; nothing is changed. " +
            "'report' measures Sent Items against them: each sent message a recipient bounced from, whether ALL its recipients failed (it reached nobody) or only some, and the bounced addresses no sent message accounts for. It counts bounces already in Deleted Items too, so it still works after a clean." +
            (readOnly ? '' :
                " 'clean' deletes the bounce-backs 'list' finds into Deleted Items, recoverable from there. Run 'list' first, show the user what would go, and get their confirmation."),
        inputSchema: z.object({
            action: z.enum(actions).describe(readOnly
                ? "'list' = the bounce-backs in the Inbox. 'report' = which sent messages bounced, and for whom."
                : "'list' = the bounce-backs in the Inbox (read-only). 'report' = which sent messages bounced, and for whom (read-only). 'clean' = delete the bounce-backs into Deleted Items."),
            email_account: z.string().describe('Outlook email account address that received the bounces (and, for report, sent the mail)'),
            days_back: z.number().int().min(1).max(3650).describe("How many days back to look (default 30). 'report' reads Sent Items over the same window, so a bounce for an older send shows as unmatched.").default(30),
        }),
        annotations: readOnly
            ? { readOnlyHint: true, idempotentHint: true, openWorldHint: false }
            : { readOnlyHint: false, destructiveHint: true, idempotentHint: false, openWorldHint: false },
    }, toolHandler(async ({ action, email_account, days_back }) => {
        if (action === 'report') {
            return json(await bridge.readBounceReport(email_account, { daysBack: days_back }));
        }
        const clean = action === 'clean';
        const result = await bridge.cleanUndeliverableEmails(email_account, { daysBack: days_back, dryRun: !clean });
        return json({
            ...result,
            note: clean
                ? 'Deleted into Deleted Items, recoverable from there until purged.'
                : "Nothing was changed. 'clean' would delete exactly these into Deleted Items.",
        });
    })));
}
