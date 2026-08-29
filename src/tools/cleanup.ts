// Destructive mailbox operations that don't belong behind move/drafts'
// narrower gates: deleting arbitrary mail by id, and permanently purging
// Deleted Items.
import { z } from 'zod';
import type { McpServer } from '@modelcontextprotocol/server';
import { deleteOutlookEmails, purgeDeletedItems } from '@dawnlit/outlook-bridge';
import { json, safe } from '../helpers';

/** Register the delete/purge tool on `server`. */
export function registerCleanupTools(server: McpServer): void {
    server.registerTool('delete_outlook_emails', {
        description:
            '⚠️ DESTRUCTIVE. Delete mail by entryId from anywhere in the account, or permanently empty Deleted Items. ' +
            "REACH FOR SOMETHING ELSE FIRST: `move_outlook_emails` to 'Deleted Items' gets mail out of the way and stays recoverable; `outlook_drafts` (action 'delete') retracts a draft behind a gate that cannot touch anything but drafts. This tool exists for the cases those two don't cover. " +
            "action 'delete' moves items to Deleted Items (so it IS recoverable), and refuses anything in the Inbox or Sent Items — INCLUDING their subfolders, since a filed folder is still received mail — unless `allow_protected` is set. " +
            '⭐ ALWAYS run with `dry_run: true` first: it resolves every id and reports the subject and folderPath it landed on WITHOUT deleting, and that is the only way to catch the failure this tool is prone to. An EntryID is an unsafe key here — Outlook silently returns a DIFFERENT message for a stale or wrong id, which is exactly what happens when several replies share one subject. Show the user the dry-run output and get explicit confirmation before the real call. ' +
            'Every result row echoes the resolved subject and folderPath, so a wrong id is visible afterwards instead of silent. ' +
            "action 'purge_deleted_items' PERMANENTLY destroys the contents of Deleted Items — nothing recovers from it. It is folder-scoped rather than id-keyed on purpose: it can only ever destroy what the operator already threw away. It also supports `dry_run`. Windows only (Outlook COM).",
        inputSchema: {
            action: z.enum(['delete', 'purge_deleted_items']).describe("'delete' = remove mail by id (to Deleted Items, recoverable); needs `entry_ids`. 'purge_deleted_items' = permanently empty Deleted Items (IRREVERSIBLE)."),
            email_account: z.string().describe('Outlook email account address'),
            entry_ids: z.array(z.string()).min(1).max(100).describe("'delete' only: EntryIDs to delete (max 100 per call). Get them from a fresh list — a stale id can resolve to a different message.").optional(),
            dry_run: z.boolean().describe('Resolve and report what WOULD happen, deleting nothing. Use it before every real call and show the user the resolved subjects/folders.').default(false),
            allow_protected: z.boolean().describe("'delete' only: permit deleting mail in the Inbox or Sent Items (or their subfolders). Default false, which refuses those and reports them. Only set this when the user has agreed to lose that specific received mail.").default(false),
            older_than_days: z.number().int().min(0).max(3650).describe("'purge_deleted_items' only: keep items newer than this many days (default 0 = purge everything).").default(0),
        },
    }, safe(async ({ action, email_account, entry_ids, dry_run, allow_protected, older_than_days }) => {
        if (action === 'purge_deleted_items') {
            const res = await purgeDeletedItems(email_account, older_than_days, dry_run);
            return json({
                ...res,
                note: dry_run
                    ? 'Dry run — nothing was purged.'
                    : 'Purged permanently. These items are NOT recoverable.',
            });
        }
        if (!entry_ids || entry_ids.length === 0) {
            throw new Error("action 'delete' needs `entry_ids` — list the mail first and pass the ids you mean.");
        }
        const res = await deleteOutlookEmails(email_account, entry_ids, {
            allowProtected: allow_protected,
            dryRun: dry_run,
        });
        return json({
            ...res,
            note: dry_run
                ? 'Dry run — nothing was deleted. Check every `subject` and `folderPath` below is the email you meant before re-running without dry_run.'
                : 'Deleted items are in Deleted Items, not destroyed — recoverable from there until purged.',
        });
    }));
}
