// Reading and organizing a mailbox: accounts, listing and searching mail, one
// email's full body, the email selected in Outlook, folder discovery, and
// filing mail into folders.
import { z } from 'zod';
import { type InboxSearchMatch, InvalidRequestError, splitQuotedOriginal } from '@dawnlit/outlook-bridge';
import { json, toolHandler } from '../results';
import type { ToolContext } from './context';

const EMAIL_ACCOUNT = z.string().describe('Outlook email account address (from get_outlook_accounts)');
const ENTRY_ID = z.string().describe('The email\'s entryId, from the listing row that found it');
const STORE_ID = z.string().describe('The storeId from the SAME listing row — it lets the email resolve in any mailbox, not just the default one').optional();

/** A subject regex from a tool argument: case-insensitive, and INVALID_REQUEST when it doesn't compile. */
function subjectRegExp(source: string | undefined): RegExp | undefined {
    if (!source) return undefined;
    try {
        return new RegExp(source, 'i');
    } catch (error) {
        throw new InvalidRequestError(`subject_regex is not a valid regular expression: ${(error as Error).message}`);
    }
}

/** A search match as the tool returns it: the body only when asked for, and capped. */
function searchRow(match: InboxSearchMatch, maxBodyChars: number | null) {
    const row = {
        entryId: match.entryId,
        storeId: match.storeId,
        subject: match.subject,
        senderName: match.senderName,
        senderEmail: match.senderEmail,
        receivedTime: match.receivedTime,
        attachmentNames: match.attachmentNames,
        folderPath: match.folderPath,
    };
    if (maxBodyChars === null) return row;
    return { ...row, body: match.body.slice(0, maxBodyChars), bodyTruncated: match.body.length > maxBodyChars };
}

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

    add('search_outlook_inbox', 'read', () => server.registerTool('search_outlook_inbox', {
        title: 'Search Outlook mail',
        description:
            "Search an account's Inbox AND every folder under it by subject, date window, attachments and folder — where list_outlook_inbox reads one folder at a time. " +
            'Use it whenever mail may already be filed into subfolders, or to find every message of a thread wherever it now sits. Sent Items, Drafts, Deleted Items and Junk are never searched. ' +
            'Returns matches newest first with entryId, storeId, subject, sender, receivedTime, attachment NAMES and folderPath; bodies only when asked for. `count` is how many matched in all.',
        inputSchema: z.object({
            email_account: EMAIL_ACCOUNT,
            days_back: z.number().int().min(1).max(3650).describe('How many days back to search (default 60). The whole tree is walked, so a wide window on a big mailbox is slow.').default(60),
            subject: z.string().max(300).describe("Subject glob, case-insensitive: `*` matches any run and `?` one character, so '*invoice*' is any subject containing it. Without a `*` the whole subject has to match.").optional(),
            subject_regex: z.string().max(500).describe("A regular expression tested against each subject, case-insensitively, for what a glob can't say ('quote|estimate'). Keep to syntax JavaScript and .NET share.").optional(),
            exclude_replies: z.boolean().describe('Leave out subjects that start with a reply or forward prefix (RE:, FW:, AW:, 回复:, …). Default false.').default(false),
            with_attachments: z.boolean().describe('Only mail with at least one attachment. Default false.').default(false),
            folders: z.array(z.string().max(300)).max(100).describe("Search only these folders, by bare name or the folderPath list_outlook_inbox_folders prints, case-insensitively. Matched exactly: naming a folder doesn't bring its subfolders, and the Inbox itself is searched only when named. Omit to search everything.").optional(),
            exclude_folders: z.array(z.string().max(300)).max(100).describe('Folders to leave out, named the same way. Each takes its whole subtree with it.').optional(),
            include_body: z.boolean().describe("Also return each match's plain-text body, capped at max_body_chars. Default false: a search is much faster without.").default(false),
            max_body_chars: z.number().int().min(100).max(20000).describe('include_body only: characters of body per match (default 2000); `bodyTruncated` says when it bit.').default(2000),
            limit: z.number().int().min(1).max(500).describe('Maximum matches returned, newest first (default 50).').default(50),
        }),
        annotations: { readOnlyHint: true, idempotentHint: true, openWorldHint: false },
    }, toolHandler(async (args) => {
        const matches = await bridge.searchInboxByFilter(args.email_account, {
            daysBack: args.days_back,
            subjectLike: args.subject,
            subjectPattern: subjectRegExp(args.subject_regex),
            excludeReplies: args.exclude_replies,
            requireAttachment: args.with_attachments,
            includeFolders: args.folders,
            excludeFolders: args.exclude_folders,
            includeBody: args.include_body,
        });
        // The walk returns folder by folder; newest first is what a reader expects.
        const newest = [...matches].sort((a, b) => b.receivedTime.localeCompare(a.receivedTime));
        return json({
            count: matches.length,
            truncated: matches.length > args.limit,
            matches: newest.slice(0, args.limit).map(match => searchRow(match, args.include_body ? args.max_body_chars : null)),
        });
    })));

    add('read_selected_outlook_email', 'read', () => server.registerTool('read_selected_outlook_email', {
        title: 'Read the selected Outlook email',
        description:
            "Read the email the user has selected, or has open, in Outlook on this machine — for 'this email' or 'the one I have open'. No account or id needed. " +
            "Returns its entryId and storeId (pass both to the other tools to act on it), subject, sender, receivedTime, attachment NAMES, and the sender's own text with the quoted thread split off. " +
            'Fails if nothing is selected or the selection is not an email.',
        inputSchema: z.object({
            max_chars: z.number().int().min(500).max(100000).describe('Cap on the returned body (default 8000); `truncated` and `bodyLength` report when it bit').default(8000),
            include_quoted: z.boolean().describe('Also return the quoted thread below the new text (default false). `quotedLength` reports its size either way.').default(false),
        }),
        annotations: { readOnlyHint: true, idempotentHint: true, openWorldHint: false },
    }, toolHandler(async ({ max_chars, include_quoted }) => {
        const email = await bridge.readSelectedEmail();
        const { body, quoted, separator } = splitQuotedOriginal(email.body);
        return json({
            entryId: email.entryId,
            storeId: email.storeId,
            subject: email.subject,
            senderName: email.senderName,
            senderEmail: email.senderEmail,
            receivedTime: email.receivedTime,
            attachmentNames: email.attachmentNames,
            body: body.slice(0, max_chars),
            truncated: body.length > max_chars,
            bodyLength: body.length,
            quoteSeparator: separator,
            quotedLength: quoted.length,
            quotedOriginal: include_quoted ? quoted.slice(0, max_chars) : '',
        });
    })));

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
