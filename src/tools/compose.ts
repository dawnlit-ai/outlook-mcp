// Composing mail: replying (including from mailbox-stored templates), fresh
// sends, signature discovery, and the template emails themselves.
import { z } from 'zod';
import { InvalidRequestError } from '@dawnlit/outlook-bridge';
import { json, text, toolHandler } from '../results';
import type { ToolContext } from './context';

/** How a template email carries several variants — stated once, referenced by the tools that use it. */
const TEMPLATE_MECHANICS =
    'A template serving several kinds of reply holds each variant between [[SECTION]] ... [[/SECTION]] markers around one shared greeting and closing, with {{PLACEHOLDER}} where text is filled in (and {{SIGNATURE}} for an Outlook signature). reply_outlook_email resolves all of that itself.';

/** What happened to a reply, as `status` has always reported it. */
function replyStatus(sendImmediately: boolean, openDraftWindow: boolean): string {
    if (sendImmediately) return 'sent';
    return openDraftWindow ? 'draft opened for review' : 'draft saved silently to the Drafts folder';
}

/** What happened to a new email, as the tool has always reported it. */
function sendStatus(sendImmediately: boolean, openDraftWindow: boolean): string {
    if (sendImmediately) return 'Email sent.';
    return openDraftWindow ? 'Email displayed in Outlook for review.' : 'Draft saved to the Outlook Drafts folder.';
}

export function registerComposeTools({ server, bridge, readOnly, add }: ToolContext): void {
    add('reply_outlook_email', 'write', () => server.registerTool('reply_outlook_email', {
        title: 'Reply to an Outlook email',
        description:
            "Reply to one email (by entry_id): a TRUE reply — your html_body above the quoted original, recipients and subject taken from the original, threaded in the recipient's mailbox. Prefer it to send_outlook_email whenever answering an email you've read. " +
            'Pass the store_id from the SAME listing row, and check the echoed repliedToSender is who you meant. ' +
            'For a templated reply, pass template_subject instead of html_body: the template\'s body, section, placeholders and signature are resolved here, so the template\'s HTML never passes through this call. ' +
            'IMPORTANT: preview the reply with the user before calling unless they approved a batch. send_immediately false + open_draft_window false saves a draft silently — the right choice for batches.',
        inputSchema: z.object({
            email_account: z.string().describe('Outlook email account address to reply as'),
            entry_id: z.string().describe('entryId of the email being replied to'),
            store_id: z.string().describe('storeId from the same listing row as entry_id').optional(),
            html_body: z.string().max(200000).describe('HTML inserted above the quoted original, used verbatim. Omit when using template_subject.').optional(),
            template_subject: z.string().max(300).describe("Reply with a template email found BY SUBJECT in the mailbox, instead of html_body. Confirm it exists first with outlook_templates (action 'list'). One of html_body / template_subject is required.").optional(),
            template_folder: z.string().max(300).describe("Folder holding the template (default 'Templates')").default('Templates'),
            template_section: z.string().max(60).describe('Which [[SECTION]] of the template to keep; the others and every marker are removed. Required when the template has sections — the call fails rather than send a body with markers in it. outlook_templates reports each template\'s sections.').optional(),
            template_placeholders: z.record(z.string(), z.string().max(50000)).describe('{{PLACEHOLDER}} → HTML to substitute, e.g. {"QUESTIONS": "line one<br>line two"}. Fails if a named placeholder isn\'t in the template.').optional(),
            signature: z.string().max(200).describe("Name of an Outlook signature (from list_outlook_signatures) to put in the template's {{SIGNATURE}} placeholder, images included. Required when the template has that placeholder.").optional(),
            send_immediately: z.boolean().describe('true sends now; false (default) stages a draft').default(false),
            open_draft_window: z.boolean().describe('Drafts only: true (default) opens a compose window; false saves silently to Drafts — use false for batches').default(true),
        }),
        annotations: { readOnlyHint: false, destructiveHint: false, idempotentHint: false, openWorldHint: true },
    }, toolHandler(async (args) => {
        if (!args.html_body && !args.template_subject) {
            throw new InvalidRequestError('Provide either html_body or template_subject.');
        }
        const result = await bridge.replyOutlookEmail({
            emailAccount: args.email_account,
            entryId: args.entry_id,
            storeId: args.store_id,
            htmlBody: args.html_body,
            templateSubject: args.template_subject,
            templateFolder: args.template_folder,
            templateSection: args.template_section,
            templatePlaceholders: args.template_placeholders,
            signatureName: args.signature,
            sendImmediately: args.send_immediately,
            openDraftWindow: args.open_draft_window,
        });
        return json({ ...result, status: replyStatus(args.send_immediately, args.open_draft_window) });
    })));

    add('send_outlook_email', 'write', () => server.registerTool('send_outlook_email', {
        title: 'Send an Outlook email',
        description:
            'Send a new email through Outlook, or stage it as a draft. IMPORTANT: before calling, show the user a preview (recipients, subject, a body summary, attachments) and get their confirmation; set send_immediately only when they explicitly ask to send without review. ' +
            'With send_immediately false (the default) the email is a DRAFT: it opens in a compose window, or — with open_draft_window false — is saved silently to Drafts, the right choice for a batch. ' +
            'To answer an email you have read, use reply_outlook_email instead.',
        inputSchema: z.object({
            email_account: z.string().describe('Outlook email account address to send from'),
            to: z.string().describe('Recipient addresses, separated by commas or semicolons'),
            cc: z.string().describe('CC addresses, separated by commas or semicolons').default(''),
            bcc: z.string().describe('BCC addresses, separated by commas or semicolons').default(''),
            subject: z.string().max(500).describe('Subject line'),
            html_body: z.string().max(200000).describe('HTML body'),
            attachment_path: z.string().max(1000).describe('Absolute path of one file to attach').optional(),
            attachment_paths: z.array(z.string().max(1000)).max(50).describe('Absolute paths of several files to attach').optional(),
            send_immediately: z.boolean().describe('true sends now; false (default) stages a draft').default(false),
            open_draft_window: z.boolean().describe('Drafts only: true (default) opens a compose window for review; false saves silently to Drafts — use false for batches').default(true),
        }),
        annotations: { readOnlyHint: false, destructiveHint: false, idempotentHint: false, openWorldHint: true },
    }, toolHandler(async (args) => {
        await bridge.sendOutlookEmail({
            emailAccount: args.email_account,
            to: args.to,
            cc: args.cc,
            bcc: args.bcc,
            subject: args.subject,
            htmlBody: args.html_body,
            attachments: [...(args.attachment_path ? [args.attachment_path] : []), ...(args.attachment_paths ?? [])],
            sendImmediately: args.send_immediately,
            openDraftWindow: args.open_draft_window,
        });
        return text(sendStatus(args.send_immediately, args.open_draft_window));
    })));

    add('list_outlook_signatures', 'read', () => server.registerTool('list_outlook_signatures', {
        title: 'List Outlook signatures',
        description:
            "List the names of the Outlook signatures set up on this machine, for reply_outlook_email's `signature`. " +
            'Exactly one name means use it; several means ASK THE USER which to sign with — signatures belong to the machine, not to an account, so never infer one from the sending address. ' +
            'An empty list means no signature is set up: ask what to sign with rather than inventing a name. Names only — the signature itself is resolved when replying.',
        inputSchema: z.object({}),
        annotations: { readOnlyHint: true, idempotentHint: true, openWorldHint: false },
    }, toolHandler(async () => {
        const names = await bridge.listOutlookSignatures();
        const note = names.length === 0
            ? 'No Outlook signatures found — ask the user what to sign with.'
            : names.length === 1
                ? 'One signature configured — use it without asking.'
                : 'Several signatures configured — ask the user which one to use.';
        return json({ signatures: names, count: names.length, note });
    })));

    const templateActions = readOnly ? (['list'] as const) : (['list', 'save'] as const);
    add('outlook_templates', 'read', () => server.registerTool('outlook_templates', {
        title: 'Outlook template emails',
        description:
            "Template emails kept in a mailbox folder (default 'Templates'): ordinary saved emails whose HTML bodies are reused as reply bodies. " +
            TEMPLATE_MECHANICS + ' ' +
            "'list' returns each template's subject, `sections` and `placeholders`. Template bodies are often Word-generated and run to tens of thousands of characters each, so the default include_body:false — subjects, sections, placeholders and a short preview — is what to use to confirm a template is intact. " +
            "Reply through reply_outlook_email's template_subject, which resolves the body itself; fetching a body (`subject` + include_body:true) is for inspecting a template that looks broken. " +
            "If the folder doesn't exist, 'list' returns folderFound:false with the mailbox's folder names — ask the user what to do rather than inventing wording. Images embedded in a template are removed from a returned htmlBody." +
            (readOnly ? '' : " 'save' adds a NEW template, creating the folder at the mailbox root if needed — only after the user has agreed; it never overwrites an existing one (templates are edited in Outlook)."),
        inputSchema: z.object({
            action: z.enum(templateActions).describe(readOnly
                ? "'list' = the folder's templates."
                : "'list' = the folder's templates (read-only). 'save' = add a new template; needs `subject` and `html_body`."),
            email_account: z.string().describe('Outlook email account address (the mailbox holding the templates)'),
            folder_name: z.string().max(300).describe("Folder holding the templates, searched case-insensitively up to 3 levels deep (default 'Templates')").default('Templates'),
            subject: z.string().max(300).describe("'list': only the template with this subject (case-insensitive) — use it whenever you need a body. 'save': the new template's subject, which is its name.").optional(),
            html_body: z.string().max(200000).describe("'save' only: the template's HTML body.").optional(),
            limit: z.number().int().min(1).max(200).describe("'list' only: maximum templates, most recently modified first (default 20)").default(20),
            include_body: z.boolean().describe("'list' only: false (default) returns subjects, sections, placeholders and a preview; true adds each full htmlBody — pair it with `subject`.").default(false),
        }),
        annotations: readOnly
            ? { readOnlyHint: true, idempotentHint: true, openWorldHint: false }
            : { readOnlyHint: false, destructiveHint: false, idempotentHint: false, openWorldHint: false },
    }, toolHandler(async ({ action, email_account, folder_name, subject, html_body, limit, include_body }) => {
        if (action === 'list') {
            return json(await bridge.readTemplateEmails(email_account, {
                folder: folder_name,
                subject,
                limit,
                includeBody: include_body
            }));
        }
        if (!subject) throw new InvalidRequestError("action 'save' needs a `subject` — it names the template.");
        if (!html_body) throw new InvalidRequestError("action 'save' needs `html_body`.");
        const result = await bridge.saveTemplateEmail(email_account, {
            subject,
            htmlBody: html_body,
            folder: folder_name
        });
        return text(
            `Template "${subject}" saved to ${result.folderPath || folder_name}`
            + (result.folderCreated ? ' (folder created)' : '')
            + '. Edit it in Outlook like any other saved email.',
        );
    })));
}
