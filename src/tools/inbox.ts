// Reading and organizing a mailbox: accounts, listing mail, one email's full
// body, folder discovery, and filing mail into folders.
import { z } from 'zod';
import { json, toolHandler } from '../results';
import type { ToolContext } from './context';

const EMAIL_ACCOUNT = z.string().describe('Outlook email account address (from get_outlook_accounts)');
const ENTRY_ID = z.string().describe('The email\'s entryId, from the listing row that found it');
const STORE_ID = z.string().describe('The storeId from the SAME listing row — it lets the email resolve in any mailbox, not just the default one').optional();

export function registerInboxTools({ server, bridge, add }: ToolContext): void {
    add('get_outlook_accounts', 'read', () => server.registerTool('get_outlook_accounts', {
        title: 'List Outlook accounts',
        description: 'List the mailboxes (as email addresses) the Outlook profile on this machine can reach. Every other Outlook tool takes one of these as `email_account`.',
        inputSchema: z.object({}),
        annotations: { readOnlyHint: true, idempotentHint: true, openWorldHint: false },
    }, toolHandler(async () => json(await bridge.getOutlookAccounts()))));

    add('list_outlook_inbox', 'read', () => server.registerTool('list_outlook_inbox', {
        title: 'List Outlook mail',
        description:
            'List recent emails from an Outlook mailbox, newest first: metadata, a short body preview, and attachment NAMES (nothing is downloaded — use read_outlook_attachment or save_outlook_attachment for those). ' +
            'By default it reads the Inbox ROOT only: mail already filed into a subfolder is invisible until you pass `folder`.',
        inputSchema: z.object({
            email_account: EMAIL_ACCOUNT,
            days_back: z.number().int().min(1).max(3650).describe('How many days back to read (default 60)').default(60),
            limit: z.number().int().min(1).max(500).describe('Maximum emails to return, newest first (default 50)').default(50),
            folder: z.string().max(300).describe(
                "Read ONE folder instead of the Inbox root. Either (a) a folder under the Inbox — a bare name ('Invoices'), a relative path ('Invoices\\2026'), or a folderPath list_outlook_inbox_folders printed; or (b) a well-known folder — 'Sent Items', 'Drafts', 'Deleted Items', 'Junk Email', 'Outbox' — optionally with a path under it ('Deleted Items\\2026'). " +
                "⭐ 'Sent Items' is how to answer \"has this already gone out?\". In Sent Items, Outbox and Drafts, `senderName`/`senderEmail` carry the RECIPIENT, since outgoing mail has no meaningful sender of its own. " +
                "A well-known name wins over a user folder of the same name — reach that one as 'Inbox\\Drafts'. Only the named folder is read, not its subfolders. Each row echoes `folderPath`; an unknown folder fails and lists the folders there are.",
            ).optional(),
            preview_chars: z.number().int().min(0).max(5000).describe('Characters of plain-text body preview per email (default 600; 0 skips bodies for a faster listing)').default(600),
        }),
        annotations: { readOnlyHint: true, idempotentHint: true, openWorldHint: false },
    }, toolHandler(async ({ email_account, days_back, limit, folder, preview_chars }) => json(
        await bridge.readInboxEmails(email_account, {
            daysBack: days_back,
            limit,
            folder,
            previewChars: preview_chars
        }),
    ))));

    add('read_outlook_email_body', 'read', () => server.registerTool('read_outlook_email_body', {
        title: 'Read an Outlook email',
        description:
            "Read ONE email's full plain-text body, when a listing's preview isn't enough. Returns the sender's own text with the quoted thread split off, plus the resolved subject and sender so you can confirm it's the email you meant. " +
            'Plain text, not HTML: an HTML table\'s rows flatten, so read tabular figures from an attachment (read_outlook_attachment) where there is one.',
        inputSchema: z.object({
            entry_id: ENTRY_ID,
            store_id: STORE_ID,
            max_chars: z.number().int().min(500).max(100000).describe('Cap on the returned body (default 8000); `truncated` and `bodyLength` report when it bit').default(8000),
            include_quoted: z.boolean().describe('Also return the quoted thread below the reply (default false). `quotedLength` reports its size either way.').default(false),
        }),
        annotations: { readOnlyHint: true, idempotentHint: true, openWorldHint: false },
    }, toolHandler(async ({ entry_id, store_id, max_chars, include_quoted }) => json(
        await bridge.readEmailBody({ entryId: entry_id, storeId: store_id }, {
            maxChars: max_chars,
            includeQuoted: include_quoted
        }),
    ))));

    add('list_outlook_inbox_folders', 'read', () => server.registerTool('list_outlook_inbox_folders', {
        title: 'List Outlook folders',
        description: "List the folders under an account's Inbox, with item counts. Use it to find where mail is filed, and before move_outlook_emails to find — and show the user — the destination.",
        inputSchema: z.object({
            email_account: EMAIL_ACCOUNT,
            max_depth: z.number().int().min(1).max(10).describe('How many levels below the Inbox to list (default 2)').default(2),
        }),
        annotations: { readOnlyHint: true, idempotentHint: true, openWorldHint: false },
    }, toolHandler(async ({ email_account, max_depth }) => json(
        await bridge.listInboxFolders(email_account, { maxDepth: max_depth }),
    ))));

    add('move_outlook_emails', 'write', () => server.registerTool('move_outlook_emails', {
        title: 'Move Outlook emails',
        description:
            "Move emails (by entryId) into a folder of the account — how processed mail gets filed. ⭐ Moving to 'Deleted Items' is also the right way to get an email out of the way: it stays recoverable there, so prefer it to delete_outlook_emails for anything but a draft. " +
            'IMPORTANT: moving someone\'s mail is a visible bulk action — find the destination with list_outlook_inbox_folders, tell the user exactly what would move where, and get their confirmation BEFORE calling. ' +
            'Moving changes each email\'s entryId, so ids saved earlier are stale afterwards. Returns the resolved folderPath, whether it was created, the count moved, and any per-email failures.',
        inputSchema: z.object({
            email_account: EMAIL_ACCOUNT,
            entry_ids: z.array(z.string()).min(1).max(100).describe('entryIds of the emails to move (max 100 per call)'),
            folder_name: z.string().max(300).describe("Destination: a folder under the Inbox by name ('Invoices') or path ('Invoices\\2026'), or a well-known folder — 'Deleted Items', 'Junk Email', 'Drafts', 'Sent Items' — optionally with a path under it. A well-known name wins over a user folder of the same name; write 'Inbox\\Drafts' for that one."),
            create_if_missing: z.boolean().describe("Create the destination when absent — the whole missing chain, so a nested path is one call (default false). Ask the user first: a mistyped path becomes a real folder instead of an error.").default(false),
        }),
        annotations: { readOnlyHint: false, destructiveHint: false, idempotentHint: false, openWorldHint: false },
    }, toolHandler(async ({ email_account, entry_ids, folder_name, create_if_missing }) => json(
        await bridge.moveOutlookEmails(email_account, entry_ids, folder_name, { createIfMissing: create_if_missing }),
    ))));
}
