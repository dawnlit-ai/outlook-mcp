// Downloading and reading Outlook attachments — a raw save-to-disk tool and
// a save-plus-extract-text (or return-as-image) tool.
import { z } from 'zod';
import type { McpServer } from '@modelcontextprotocol/sdk/server/mcp.js';
import { saveEmailAttachment, saveEmailAttachmentDetailed } from '@dawnlit/outlook-bridge';
import { contents, json, safe, text } from '../helpers';
import { extractAttachmentText } from '../attachmentText';

/** Register the attachment-saving and attachment-reading tools on `server`. */
export function registerAttachmentTools(server: McpServer): void {
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
}
