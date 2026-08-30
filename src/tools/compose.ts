// Composing and sending mail: replying (including from mailbox-stored
// templates), fresh sends, signature discovery, and template management.
import { z } from 'zod';
import type { McpServer } from '@modelcontextprotocol/server';
import {
    listOutlookSignatures,
    readTemplateEmails,
    replyOutlookEmail,
    saveTemplateEmail,
    sendOutlookEmail,
} from '@dawnlit/outlook-bridge';
import { json, safe, text } from '../helpers';

/** How a template email carries several reply variants. Stated once here and
 *  referenced (not restated) by reply_outlook_email's template params. */
const TEMPLATE_MECHANICS =
    'A template that serves several reply types holds each variant between [[SECTION]] ... [[/SECTION]] markers around one shared greeting/closing/signature, with {{PLACEHOLDER}} where the caller fills text in (and {{SIGNATURE}} for an Outlook signature). reply_outlook_email resolves all of that server-side.';

/** Register reply, send, signature-listing, and template tools on `server`. */
export function registerComposeTools(server: McpServer): void {
    server.registerTool('reply_outlook_email', {
        description: "Reply to a specific email (by entry_id from list_outlook_inbox): creates a TRUE reply — your html_body inserted above the quoted original, To/subject taken from the original, threads correctly in the recipient's mailbox. Prefer this over send_outlook_email when answering an email you've read; use send_outlook_email only for fresh outreach. Pass the store_id from the SAME listing row, and verify the echoed repliedToSender matches who you meant (same-thread rows have look-alike EntryIDs). For a templated reply, pass template_subject instead of html_body — the server resolves the saved template's body, section, placeholders and signature itself, so the (often large) template HTML never crosses this call and there is no size cap to hit. IMPORTANT: preview with the user before calling unless they've approved a batch. Draft mode mirrors send_outlook_email: send_immediately false + open_draft_window false saves silently to Drafts — the right choice for batches.",
        inputSchema: z.object({
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
        }),
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
        inputSchema: z.object({}),
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
        inputSchema: z.object({
            action: z.enum(['list', 'save']).describe("'list' = the folder's templates (read-only). 'save' = add a new template; needs `subject` and `html_body`."),
            email_account: z.string().describe('Outlook email account address (the mailbox holding the Templates folder)'),
            folder_name: z.string().max(200).describe("Mailbox folder holding the templates (default 'Templates'; searched case-insensitively up to 3 levels deep, and for 'save' created at the mailbox root if absent)").default('Templates'),
            subject: z.string().max(300).describe("'list': return only the template with this exact subject (case-insensitive) — use it whenever you need a body, since fetching one is safe and fetching a folder's worth is what overruns the limit. 'save': the new template's subject, which serves as its name.").optional(),
            html_body: z.string().describe("'save' only: the template's HTML body, reused later as a reply body.").optional(),
            limit: z.number().int().min(1).max(50).describe("'list' only: maximum templates to return, most recently modified first (default 20)").default(20),
            include_body: z.boolean().describe("'list' only: false (default) returns subjects, sections, placeholders and a short bodyPreview; true also returns each full htmlBody — pair it with `subject` so only one body is pulled.").default(false),
        }),
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

    server.registerTool('send_outlook_email', {
        description: 'Send an email via Outlook, or stage it as a draft. IMPORTANT: Before calling this tool, always show the user a preview of the email (recipients, subject, body summary) and ask for confirmation. Only set send_immediately to true if the user explicitly asks to send without preview. When send_immediately is false (default), the email is a DRAFT: by default it opens in an Outlook compose window for review; set open_draft_window to false to save it silently to the Drafts folder with NO window — the right choice for a batch, so you don\'t pop dozens of windows. Silent drafts land in Outlook\'s Drafts folder for the user to review and send (send_all_drafts can send them in bulk once approved).',
        inputSchema: z.object({
            email_account: z.string().describe('Outlook email account address (sender)'),
            to: z.string().describe('Recipient email addresses (semicolon-separated)'),
            cc: z.string().describe('CC email addresses (semicolon-separated)').default(''),
            subject: z.string().max(500).describe('Email subject line'),
            html_body: z.string().max(100000).describe('Email HTML body content'),
            attachment_path: z.string().max(1000).describe('Absolute path to attachment file').optional(),
            send_immediately: z.boolean().describe('If true, send immediately; if false, stage as a draft').default(false),
            open_draft_window: z.boolean().describe('Draft mode only (send_immediately false): true (default) opens a compose window for review; false saves silently to the Drafts folder with no window — use false for batches so you don\'t open a window per email').default(true),
        }),
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
}
