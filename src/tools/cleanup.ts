// Destructive mailbox operations that don't belong behind move or drafts'
// narrower gates: deleting mail by id, and permanently emptying Deleted Items.
import { z } from 'zod';
import { InvalidRequestError } from '@dawnlit/outlook-bridge';
import { json, toolHandler } from '../results';
import type { ToolContext } from './context';

export function registerCleanupTools({ server, bridge, add }: ToolContext): void {
    add('delete_outlook_emails', 'write', () => server.registerTool('delete_outlook_emails', {
        title: 'Delete Outlook emails',
        description:
            '⚠️ DESTRUCTIVE. Delete mail by entryId from anywhere in an account, or permanently empty Deleted Items. ' +
            "REACH FOR SOMETHING ELSE FIRST: move_outlook_emails to 'Deleted Items' gets mail out of the way recoverably, and outlook_drafts (action 'delete') retracts drafts behind a gate that can touch nothing else. " +
            "action 'delete' moves items to Deleted Items (recoverable) and refuses anything in the Inbox or Sent Items — INCLUDING their subfolders, since filed mail is still received mail — unless `allow_protected` is set. " +
            '⭐ ALWAYS run with `dry_run: true` first: it resolves every id and reports the subject and folder it landed on WITHOUT deleting. An entryId is an unsafe key — a stale or wrong id can resolve to a DIFFERENT email, which is easy to do when replies share a subject. Show the user the dry-run output and get their confirmation before the real call. ' +
            "action 'purge_deleted_items' PERMANENTLY destroys the contents of Deleted Items — nothing recovers from it. It can only ever destroy what was already thrown away, and supports `dry_run` too.",
        inputSchema: z.object({
            action: z.enum(['delete', 'purge_deleted_items']).describe("'delete' = remove mail by id into Deleted Items (recoverable); needs `entry_ids`. 'purge_deleted_items' = permanently empty Deleted Items (IRREVERSIBLE)."),
            email_account: z.string().describe('Outlook email account address'),
            entry_ids: z.array(z.string()).min(1).max(100).describe("'delete' only: entryIds to delete (max 100 per call). Take them from a fresh listing — a stale id can resolve to a different email.").optional(),
            dry_run: z.boolean().describe('Resolve and report what WOULD happen, deleting nothing. Run it before every real call and show the user the result.').default(false),
            allow_protected: z.boolean().describe("'delete' only: allow deleting mail in the Inbox or Sent Items (or their subfolders). Set it only when the user has agreed to lose that specific mail.").default(false),
            older_than_days: z.number().int().min(0).max(36500).describe("'purge_deleted_items' only: keep items newer than this many days (default 0 = purge everything).").default(0),
        }),
        annotations: { readOnlyHint: false, destructiveHint: true, idempotentHint: false, openWorldHint: false },
    }, toolHandler(async ({ action, email_account, entry_ids, dry_run, allow_protected, older_than_days }) => {
        if (action === 'purge_deleted_items') {
            const result = await bridge.purgeDeletedItems(email_account, {
                olderThanDays: older_than_days,
                dryRun: dry_run
            });
            return json({
                ...result,
                note: dry_run ? 'Dry run — nothing was purged.' : 'Purged permanently. These items are NOT recoverable.'
            });
        }
        if (!entry_ids || entry_ids.length === 0) {
            throw new InvalidRequestError("action 'delete' needs `entry_ids` — list the mail first and pass the ids you mean.");
        }
        const result = await bridge.deleteOutlookEmails(email_account, entry_ids, {
            allowProtected: allow_protected,
            dryRun: dry_run
        });
        return json({
            ...result,
            note: dry_run
                ? 'Dry run — nothing was deleted. Check every `subject` and `folderPath` below is the email you meant before running without dry_run.'
                : 'Deleted items are in Deleted Items, recoverable from there until purged.',
        });
    })));
}
