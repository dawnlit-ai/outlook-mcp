// Reading and organizing an Outlook mailbox: accounts, inbox listing, a
// single email's full body, folder discovery, and moving mail between
// folders.
import { z } from 'zod';
import type { McpServer } from '@modelcontextprotocol/server';
import {
    getOutlookAccounts,
    listInboxFolders,
    moveOutlookEmails,
    readEmailBody,
    readInboxEmails,
} from '@dawnlit/outlook-bridge';
import { json, safe } from '../helpers';

/** Register account listing, inbox reading, folder listing, and move tools on `server`. */
export function registerInboxTools(server: McpServer): void {
    server.registerTool('get_outlook_accounts', {
        description: 'Get list of Outlook email accounts configured on this system',
    }, safe(async () => {
        const accounts = await getOutlookAccounts();
        return json(accounts);
    }));

    server.registerTool('list_outlook_inbox', {
        description: "List recent emails from an Outlook inbox. Returns metadata and a short body preview — attachments are NOT downloaded (use save_outlook_attachment / read_outlook_attachment for those). By default it reads the Inbox ROOT only: mail already filed into a subfolder is invisible until you pass `folder`.",
        inputSchema: {
            email_account: z.string().describe('Outlook email account address'),
            days_back: z.number().int().min(1).max(365).describe('How many days back to scan (default 60)').default(60),
            limit: z.number().int().min(1).max(200).describe('Maximum number of emails to return, newest first (default 50)').default(50),
            folder: z.string().max(300).describe("Read ONE folder instead of the Inbox root. Two kinds of value: (a) a folder under the Inbox — a bare name ('Invoices'), a relative path ('Invoices\\2026'), or the full folderPath list_outlook_inbox_folders prints; (b) a WELL-KNOWN folder by name — 'Sent Items', 'Drafts', 'Deleted Items', 'Junk Email', 'Outbox' — optionally with a path under it ('Deleted Items\\2026'). ⭐ 'Sent Items' is how you answer \"has this already gone out?\". Sent Items/Outbox/Drafts have no ReceivedTime, so they are filtered and stamped on SentOn / LastModificationTime instead, and `senderName`/`senderEmail` carry the RECIPIENT there (outgoing mail has no meaningful sender of its own). A well-known name wins over a user folder of the same name — reach that one as 'Inbox\\Drafts'. Only the named folder is read, NOT its subfolders. Each row echoes `folderPath` so you can confirm what matched; unknown names throw and list the alternatives rather than silently falling back to the root.").optional(),
        },
    }, safe(async ({ email_account, days_back, limit, folder }) => {
        const entries = await readInboxEmails(email_account, days_back, limit, folder);
        return json(entries);
    }));

    server.registerTool('read_outlook_email_body', {
        description: "Read ONE email's full plain-text body by entryId, when list_outlook_inbox's ~600-char bodyPreview isn't enough. Returns the sender's own text with the quoted thread stripped off (and the resolved subject/sender so you can confirm it's the email you meant). Plain text, not HTML: an HTML table's rows flatten, so read a tabular figure from its attachment (read_outlook_attachment) instead. Pass the store_id from the same list_outlook_inbox row so the email resolves unambiguously across mailboxes.",
        inputSchema: {
            entry_id: z.string().describe('Outlook email EntryID (from list_outlook_inbox)'),
            store_id: z.string().describe('Outlook StoreID from the same list_outlook_inbox row — disambiguate the email across mailboxes').optional(),
            max_chars: z.number().int().min(500).max(50000).describe('Cap on the returned body (default 8000); `truncated` and `bodyLength` report when it bit').default(8000),
            include_quoted: z.boolean().describe('Also return the quoted thread below the reply (default false). Set true only to mine the thread for the original message. `quotedLength` reports its size either way.').default(false),
        },
    }, safe(async ({ entry_id, store_id, max_chars, include_quoted }) => {
        const result = await readEmailBody(entry_id, store_id, max_chars, include_quoted);
        return json(result);
    }));

    server.registerTool('list_outlook_inbox_folders', {
        description: "List the folders under an Outlook account's Inbox, with item counts. Use before move_outlook_emails to find the right destination and present the choices to the user.",
        inputSchema: {
            email_account: z.string().describe('Outlook email account address'),
            max_depth: z.number().int().min(1).max(4).describe('How many levels below Inbox to list (default 2)').default(2),
        },
    }, safe(async ({ email_account, max_depth }) => {
        const folders = await listInboxFolders(email_account, max_depth);
        return json(folders);
    }));

    server.registerTool('move_outlook_emails', {
        description: "Move emails (by entryId from list_outlook_inbox) into any folder of the account — used to file processed emails. ⭐ Moving to 'Deleted Items' is ALSO how you delete mail reversibly, and is the right way to get an email out of the way: it survives in Deleted Items instead of being destroyed, so prefer it over delete_outlook_emails for anything that isn't a draft. IMPORTANT: moving the user's mail is a visible bulk action — list the folders (list_outlook_inbox_folders), tell the user exactly what would move where, and get explicit confirmation BEFORE calling. Moving changes each item's EntryID, so any saved ids are stale afterwards. Returns the resolved folderPath, whether it was created, the moved count and per-item failures.",
        inputSchema: {
            email_account: z.string().describe('Outlook email account address'),
            entry_ids: z.array(z.string()).min(1).max(100).describe('EntryIDs of the emails to move (from list_outlook_inbox; max 100 per call)'),
            folder_name: z.string().max(200).describe("Destination: a folder under the Inbox by bare name ('Invoices') or nested path ('Invoices\\2026'), or a well-known folder — 'Deleted Items', 'Junk Email', 'Drafts', 'Sent Items' — optionally with a path under it. A well-known name wins over a user folder of the same name; qualify as 'Inbox\\Drafts' to reach that one."),
            create_if_missing: z.boolean().describe('Create the destination if absent — the WHOLE missing chain, so a nested path is one call rather than needing the parent made by hand first (default false; ask the user first, and note that a typo\'d path then becomes a real folder rather than an error)').default(false),
        },
    }, safe(async ({ email_account, entry_ids, folder_name, create_if_missing }) => {
        const result = await moveOutlookEmails(email_account, entry_ids, folder_name, create_if_missing);
        return json(result);
    }));
}
