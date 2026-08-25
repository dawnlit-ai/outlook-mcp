// Generic Outlook MCP tools: everything @dawnlit/outlook-bridge can do, exposed
// as MCP tools any host server can register. Nothing here knows about any
// particular application's domain — no email templates beyond what a mailbox
// itself stores, no business rules about what counts as a given kind of email.
// A host that wants those layers composes them on top by registering its own
// tools alongside these.
import { z } from 'zod';
import type { McpServer } from '@modelcontextprotocol/sdk/server/mcp.js';
import {
    deleteOutlookDrafts,
    deleteOutlookEmails,
    getOutlookAccounts,
    listInboxFolders,
    listOutlookDrafts,
    listOutlookSignatures,
    moveOutlookEmails,
    purgeDeletedItems,
    readEmailBody,
    readInboxEmails,
    readTemplateEmails,
    replyOutlookEmail,
    saveEmailAttachment,
    saveEmailAttachmentDetailed,
    saveTemplateEmail,
    sendAllDrafts,
    sendDrafts,
    sendOutlookEmail,
} from '@dawnlit/outlook-bridge';
import { contents, json, safe, text } from './helpers';
import { extractAttachmentText } from './attachmentText';

/** How a template email carries several reply variants. Stated once here and
 *  referenced (not restated) by reply_outlook_email's template params. */
const TEMPLATE_MECHANICS =
    'A template that serves several reply types holds each variant between [[SECTION]] ... [[/SECTION]] markers around one shared greeting/closing/signature, with {{PLACEHOLDER}} where the caller fills text in (and {{SIGNATURE}} for an Outlook signature). reply_outlook_email resolves all of that server-side.';

/**
 * Register every generic Outlook tool — accounts, inbox reading, folders,
 * replying/sending, drafts, signatures, mailbox-stored templates, attachments,
 * and mailbox cleanup — on `server`. Every tool is backed purely by
 * `@dawnlit/outlook-bridge`; none of it depends on anything the host app owns.
 */
export function registerOutlookTools(server: McpServer): void {
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

    server.registerTool('reply_outlook_email', {
        description: "Reply to a specific email (by entry_id from list_outlook_inbox): creates a TRUE reply — your html_body inserted above the quoted original, To/subject taken from the original, threads correctly in the recipient's mailbox. Prefer this over send_outlook_email when answering an email you've read; use send_outlook_email only for fresh outreach. Pass the store_id from the SAME listing row, and verify the echoed repliedToSender matches who you meant (same-thread rows have look-alike EntryIDs). For a templated reply, pass template_subject instead of html_body — the server resolves the saved template's body, section, placeholders and signature itself, so the (often large) template HTML never crosses this call and there is no size cap to hit. IMPORTANT: preview with the user before calling unless they've approved a batch. Draft mode mirrors send_outlook_email: send_immediately false + open_draft_window false saves silently to Drafts — the right choice for batches.",
        inputSchema: {
            email_account: z.string().describe('Outlook email account address to send/draft the reply as'),
            entry_id: z.string().describe('EntryID of the email being replied to (from list_outlook_inbox)'),
            store_id: z.string().describe('Outlook StoreID from the same list_outlook_inbox row — disambiguate the email across mailboxes').optional(),
            html_body: z.string().max(100000).describe('HTML inserted above the quoted original, used verbatim. Omit when using template_subject.').optional(),
            template_subject: z.string().max(300).describe('Reply with a template email resolved BY SUBJECT from the mailbox instead of passing html_body. Confirm the subject exists first with outlook_templates (action \'list\', include_body:false is enough). One of html_body / template_subject is required.').optional(),
            template_folder: z.string().max(200).describe("Folder holding the template when template_subject is used (default 'Templates')").default('Templates'),
            template_section: z.string().max(60).describe('Which [[SECTION]] of the template to keep; the others and all markers are stripped server-side. Required if the template has any — the call fails rather than mail out a body with markers in it. outlook_templates reports each template\'s sections.').optional(),
            template_placeholders: z.record(z.string(), z.string().max(20000)).describe('{{PLACEHOLDER}} → HTML to substitute, e.g. {"QUESTIONS": "line one<br>line two"}. Tolerant of Word splitting a placeholder across tags. Fails if the named placeholder isn\'t in the template.').optional(),
            signature: z.string().max(200).describe("Name of an Outlook signature (from list_outlook_signatures) to substitute into the template's {{SIGNATURE}} placeholder — resolved server-side, images included. Required when the template has that placeholder; the call fails rather than mail out an unsigned reply. Omit for templates whose signature is written into the body.").optional(),
            send_immediately: z.boolean().describe('If true, send immediately; if false (default), stage as a draft').default(false),
            open_draft_window: z.boolean().describe('Draft mode only: true (default) opens a compose window; false saves silently to the Drafts folder — use false for batches').default(true),
        },
    }, safe(async ({
                       email_account,
                       entry_id,
                       store_id,
                       html_body,
                       template_subject,
                       template_folder,
                       template_section,
                       template_placeholders,
                       signature,
                       send_immediately,
                       open_draft_window,
                   }) => {
        if (!html_body && !template_subject) {
            throw new Error('Provide either html_body or template_subject.');
        }
        const result = await replyOutlookEmail({
            emailAccount: email_account,
            entryId: entry_id,
            storeId: store_id,
            htmlBody: html_body,
            templateSubject: template_subject,
            templateFolder: template_folder,
            templateSection: template_section,
            templatePlaceholders: template_placeholders,
            signatureName: signature,
            sendImmediately: send_immediately,
            openDraftWindow: open_draft_window,
        });
        return json({
            ...result,
            status: send_immediately ? 'sent' : (open_draft_window ? 'draft opened for review' : 'draft saved silently to the Drafts folder'),
        });
    }));

    server.registerTool('list_outlook_signatures', {
        description: "List the names of the Outlook signatures configured on this machine (the .htm files in the user's Signatures folder). Use before reply_outlook_email's `signature` when the template holds a {{SIGNATURE}} placeholder: exactly one name means use it, several means ASK THE USER which to sign with — the names are machine-wide and carry no account association, so never infer one from the sending address. An empty list means no signature is set up; ask the user what to sign with rather than inventing a name or sending unsigned. Returns names only — the signature HTML is resolved server-side when replying and never enters the conversation.",
        inputSchema: {},
    }, safe(async () => {
        const names = await listOutlookSignatures();
        return json({
            signatures: names,
            count: names.length,
            note: names.length === 0
                ? 'No Outlook signatures found — ask the user what to sign with.'
                : names.length === 1
                    ? 'One signature configured — use it without asking.'
                    : 'Several signatures configured — ask the user which one to use.',
        });
    }));

    server.registerTool('outlook_templates', {
        description:
            "Template emails kept in an Outlook mailbox folder (default 'Templates') — ordinary saved emails whose HTML bodies are reused as reply bodies. " +
            TEMPLATE_MECHANICS + ' ' +
            "'list' returns each template's subject, `sections` and `placeholders`. ⚠️ These bodies are often Word-generated and can run tens of thousands of characters EACH, so include_body:true across a folder can blow the response limit — the default include_body:false gives subjects, sections, placeholders and a short preview, which is everything you need to confirm a template is intact before drafting from it. Always reply via reply_outlook_email's template_subject, which resolves the body server-side; fetching a body with `subject` + include_body:true is a last resort, for inspecting a template that looks broken. If the folder doesn't exist, 'list' returns folderFound:false plus the mailbox's folder names — ASK THE USER what to do rather than silently inventing wording. Images embedded in a template (signature logos etc.) are STRIPPED from a returned htmlBody, since neither cid: nor file: refs survive into a new email reliably. " +
            "'save' adds a NEW template, creating the folder at the mailbox root if missing — only after the user has explicitly agreed to create it, and never to 'update' an existing one (it always adds, never overwrites; the user edits templates in Outlook).",
        inputSchema: {
            action: z.enum(['list', 'save']).describe("'list' = the folder's templates (read-only). 'save' = add a new template; needs `subject` and `html_body`."),
            email_account: z.string().describe('Outlook email account address (the mailbox holding the Templates folder)'),
            folder_name: z.string().max(200).describe("Mailbox folder holding the templates (default 'Templates'; searched case-insensitively up to 3 levels deep, and for 'save' created at the mailbox root if absent)").default('Templates'),
            subject: z.string().max(300).describe("'list': return only the template with this exact subject (case-insensitive) — use it whenever you need a body, since fetching one is safe and fetching a folder's worth is what overruns the limit. 'save': the new template's subject, which serves as its name.").optional(),
            html_body: z.string().describe("'save' only: the template's HTML body, reused later as a reply body.").optional(),
            limit: z.number().int().min(1).max(50).describe("'list' only: maximum templates to return, most recently modified first (default 20)").default(20),
            include_body: z.boolean().describe("'list' only: false (default) returns subjects, sections, placeholders and a short bodyPreview; true also returns each full htmlBody — pair it with `subject` so only one body is pulled.").default(false),
        },
    }, safe(async ({ action, email_account, folder_name, subject, html_body, limit, include_body }) => {
        if (action === 'list') {
            return json(await readTemplateEmails(email_account, folder_name, limit, include_body, subject));
        }
        if (!subject) throw new Error("action 'save' needs a `subject` — it names the template.");
        if (!html_body) throw new Error("action 'save' needs `html_body`.");
        const result = await saveTemplateEmail(email_account, subject, html_body, folder_name);
        return text(
            `Template "${subject}" saved to ${result.folderPath || folder_name}`
            + (result.folderCreated ? ' (folder created)' : '')
            + '. Edit it in Outlook like any other saved email.',
        );
    }));

    server.registerTool('save_outlook_attachment', {
        description: "Download a specific attachment from an Outlook email to a temp directory, returning the saved file path. Pass the entryId (and, when the same row has one, the store_id) so the email resolves unambiguously across multiple mailboxes.",
        inputSchema: {
            entry_id: z.string().describe('Outlook email EntryID (from list_outlook_inbox)'),
            file_name: z.string().max(500).describe("Exact attachment filename to save (e.g. 'invoice.pdf')"),
            store_id: z.string().describe('Outlook StoreID from the same list_outlook_inbox row — disambiguate the email across mailboxes').optional(),
        },
    }, safe(async ({ entry_id, file_name, store_id }) => {
        const savedPath = await saveEmailAttachment(entry_id, file_name, store_id);
        return text(savedPath);
    }));

    server.registerTool('read_outlook_attachment', {
        description: "Open an attachment on an Outlook email and return its contents. Pass the entry_id AND the store_id from the SAME list_outlook_inbox row (store_id disambiguate the email across mailboxes), plus the exact attachment file name (from that email's attachmentNames). Extracts text from PDF, Excel (.xlsx), and plain-text/CSV attachments. Image attachments (a scanned or screenshot document) are returned as an inline image alongside the metadata, so read it directly — there's no OCR step to ask for. The result also echoes the resolved email's subject + senderEmail: verify these match who you intended before relying on the numbers, since many replies in a thread share one subject and it's easy to pass a neighboring row's entry_id.",
        inputSchema: {
            entry_id: z.string().describe('Outlook email EntryID (from list_outlook_inbox)'),
            file_name: z.string().max(500).describe('Exact attachment filename to open (from the email\'s attachmentNames)'),
            store_id: z.string().describe('Outlook StoreID from the same list_outlook_inbox row — disambiguate the email across mailboxes').optional(),
            max_chars: z.number().int().min(500).max(100000).describe('Max characters of extracted text to return (default 20000)').default(20000),
        },
    }, safe(async ({ entry_id, file_name, store_id, max_chars }) => {
        const saved = await saveEmailAttachmentDetailed(entry_id, file_name, store_id);
        const extracted = await extractAttachmentText(saved.path);
        const base = { file: saved.path, subject: saved.subject, senderEmail: saved.senderEmail, type: extracted.type };

        if ('data' in extracted) {
            return contents(
                { type: 'text', text: JSON.stringify({ ...base, mimeType: extracted.mimeType }, null, 2) },
                extracted,
            );
        }

        if (!extracted.text) {
            return json({ ...base, text: '', note: extracted.note });
        }

        const truncated = extracted.text.length > max_chars;
        return json({
            ...base,
            chars: extracted.text.length,
            truncated,
            text: truncated ? extracted.text.slice(0, max_chars) : extracted.text,
        });
    }));

    server.registerTool('send_outlook_email', {
        description: 'Send an email via Outlook, or stage it as a draft. IMPORTANT: Before calling this tool, always show the user a preview of the email (recipients, subject, body summary) and ask for confirmation. Only set send_immediately to true if the user explicitly asks to send without preview. When send_immediately is false (default), the email is a DRAFT: by default it opens in an Outlook compose window for review; set open_draft_window to false to save it silently to the Drafts folder with NO window — the right choice for a batch, so you don\'t pop dozens of windows. Silent drafts land in Outlook\'s Drafts folder for the user to review and send (send_all_drafts can send them in bulk once approved).',
        inputSchema: {
            email_account: z.string().describe('Outlook email account address (sender)'),
            to: z.string().describe('Recipient email addresses (semicolon-separated)'),
            cc: z.string().describe('CC email addresses (semicolon-separated)').default(''),
            subject: z.string().max(500).describe('Email subject line'),
            html_body: z.string().max(100000).describe('Email HTML body content'),
            attachment_path: z.string().max(1000).describe('Absolute path to attachment file').optional(),
            send_immediately: z.boolean().describe('If true, send immediately; if false, stage as a draft').default(false),
            open_draft_window: z.boolean().describe('Draft mode only (send_immediately false): true (default) opens a compose window for review; false saves silently to the Drafts folder with no window — use false for batches so you don\'t open a window per email').default(true),
        },
    }, safe(async ({
                       email_account,
                       to,
                       cc,
                       subject,
                       html_body,
                       attachment_path,
                       send_immediately,
                       open_draft_window,
                   }) => {
        await sendOutlookEmail({
            emailAccount: email_account,
            to,
            cc,
            subject,
            htmlBody: html_body,
            attachmentPath: attachment_path,
            sendImmediately: send_immediately,
            openDraftWindow: open_draft_window,
        });

        return text(send_immediately
            ? 'Email sent.'
            : open_draft_window
                ? 'Email displayed in Outlook for review.'
                : 'Draft saved to the Outlook Drafts folder.');
    }));

    server.registerTool('outlook_drafts', {
        description:
            "The mail drafts staged for an account — what is actually sitting in Outlook's Drafts folder(s), and the ways to remove or send one. " +
            "'list' is READ-ONLY and is the only way to see what a batch of staged mail really contains: call it BEFORE send_all_drafts, which is account-wide and sends everything it finds. " +
            "It answers the two questions a draft-staging run cannot answer for itself — whether an earlier batch is still pending or was already sent, and whether re-drafting has left TWO drafts for one recipient. " +
            'Each row gives entryId, subject, To, a short bodyPreview and the folder, newest first; the preview is what distinguishes one templated variant from another. Full bodies are never returned — a templated body can run tens of thousands of characters. ' +
            "'delete' removes drafts by entryId. Outlook moves them to Deleted Items, so a mistake is recoverable there, and every id is verified to be a mail draft belonging to THIS account before anything happens — an id pointing at ordinary inbox mail is refused, not deleted. " +
            "Use it to retract superseded drafts: re-drafting a reply does NOT replace the earlier one, it adds a second, and bulk-sending would then send both. " +
            "'send' sends only the drafts named by `entry_ids` — the reviewed subset from a 'list' call — instead of send_all_drafts' account-wide sweep. Prefer this whenever the user wants to see what would go out and pick some of it, rather than flush the whole folder. Same ownership gate as 'delete': an id pointing at ordinary mail, or at another account's draft, is refused and reported rather than sent. " +
            '⚠️ Deleting or sending the user\'s mail is a visible, hard-to-undo action — list first, tell them exactly which drafts would go, and get explicit confirmation. Never delete or send a draft the user has not agreed to. ' +
            'Drafts bound to a different account are invisible to every action. Windows only (Outlook COM).',
        inputSchema: {
            action: z.enum(['list', 'delete', 'send']).describe("'list' = the account's drafts (read-only). 'delete' = remove drafts by id; needs `entry_ids`. 'send' = send only the named drafts; needs `entry_ids`."),
            email_account: z.string().describe('Outlook email account address whose drafts to list, delete, or send'),
            entry_ids: z.array(z.string()).min(1).max(100).describe("'delete'/'send' only: EntryIDs from a preceding 'list' (max 100 per call)").optional(),
            limit: z.number().int().min(1).max(500).describe("'list' only: maximum drafts to return, newest first (default 100); `count` reports the true total and `truncated` says whether it bit").default(100),
            preview_chars: z.number().int().min(0).max(2000).describe("'list' only: characters of plain-text body preview per draft (default 300). Raise it when several drafts share a greeting and you need to read further in to tell them apart.").default(300),
        },
    }, safe(async ({ action, email_account, entry_ids, limit, preview_chars }) => {
        if (action === 'list') {
            return json(await listOutlookDrafts(email_account, limit, preview_chars));
        }
        if (!entry_ids || entry_ids.length === 0) {
            throw new Error(`action '${action}' needs \`entry_ids\` — list the drafts first and pass the ids you mean.`);
        }
        if (action === 'send') {
            const res = await sendDrafts(email_account, entry_ids);
            return json(res);
        }
        const res = await deleteOutlookDrafts(email_account, entry_ids);
        return json({
            ...res,
            note: "Deleted drafts are in Outlook's Deleted Items folder, not gone — recoverable from there.",
        });
    }));

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

    server.registerTool('send_all_drafts', {
        description: "Send every mail draft in Outlook's Drafts folder that is set to send from the given account. Each matching draft goes out IMMEDIATELY, so this is a bulk send: before calling, tell the user it will send all queued drafts for that account at once and get explicit confirmation. It is batch-blind — it cannot send only a batch you just staged. Call `outlook_drafts` (action 'list') FIRST and reconcile what's there against what you staged: a count higher than yours means unrelated or superseded drafts would go out too, and superseded ones should be removed with `outlook_drafts` (action 'delete') before sending. Prefer `outlook_drafts` (action 'send') instead whenever the user wants to review the list and pick which ones go — it sends only the entry_ids you give it. Drafts bound to a different account, or with no sending account set, are left untouched. Returns how many were sent and any that failed (by subject). Use get_outlook_accounts first if you're unsure of the exact account address.",
        inputSchema: {
            email_account: z.string().describe('Outlook email account address (sender) whose drafts should be sent'),
        },
    }, safe(async ({ email_account }) => {
        const res = await sendAllDrafts(email_account);
        return json(res);
    }));
}
