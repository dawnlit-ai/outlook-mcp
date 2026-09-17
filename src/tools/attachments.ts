// Attachments: saving one to disk, and reading one — extracted text, or the
// image itself.
import { z } from 'zod';
import { extractAttachmentContent } from '../attachmentContent';
import { contents, json, text, toolHandler } from '../results';
import type { ToolContext } from './context';

const ENTRY_ID = z.string().describe('The email\'s entryId, from the listing row that found it');
const STORE_ID = z.string().describe('The storeId from the SAME listing row — it lets the email resolve in any mailbox, not just the default one').optional();

// Both tools leave the mailbox as it is; they only write a copy of the file to a
// private scratch directory, which is why they count as reading tools.
export function registerAttachmentTools({ server, bridge, add }: ToolContext): void {
    add('save_outlook_attachment', 'read', () => server.registerTool('save_outlook_attachment', {
        title: 'Save an Outlook attachment',
        description: 'Save one attachment of an email to a private temporary directory and return the saved file\'s path. Pass the entry_id and store_id from the same listing row, and the file name exactly as the email\'s attachmentNames lists it.',
        inputSchema: z.object({
            entry_id: ENTRY_ID,
            file_name: z.string().max(500).describe("The attachment's file name (e.g. 'invoice.pdf'), from the email's attachmentNames"),
            store_id: STORE_ID,
        }),
        annotations: { readOnlyHint: true, idempotentHint: true, openWorldHint: false },
    }, toolHandler(async ({ entry_id, file_name, store_id }) => {
        const saved = await bridge.saveEmailAttachment({ entryId: entry_id, storeId: store_id }, file_name);
        return text(saved.path);
    })));

    add('read_outlook_attachment', 'read', () => server.registerTool('read_outlook_attachment', {
        title: 'Read an Outlook attachment',
        description:
            "Open one attachment of an email and return its contents: the text of a PDF, a spreadsheet (.xlsx, one tab-separated row per line) or a text file; a JPEG, PNG, GIF or WebP image as an inline image to look at directly. Another format, or an image too large to inline, comes back as a note with the saved file's path. " +
            'Pass the entry_id and store_id from the same listing row, and the file name from its attachmentNames. The result echoes the resolved email\'s subject and senderEmail: check they match the email you meant before relying on what you read, since replies in one thread share a subject.',
        inputSchema: z.object({
            entry_id: ENTRY_ID,
            file_name: z.string().max(500).describe("The attachment's file name, from the email's attachmentNames"),
            store_id: STORE_ID,
            max_chars: z.number().int().min(500).max(200000).describe('Maximum characters of extracted text to return (default 20000)').default(20000),
        }),
        annotations: { readOnlyHint: true, idempotentHint: true, openWorldHint: false },
    }, toolHandler(async ({ entry_id, file_name, store_id, max_chars }) => {
        const saved = await bridge.saveEmailAttachment({ entryId: entry_id, storeId: store_id }, file_name);
        const extracted = await extractAttachmentContent(saved.path);
        const about = { file: saved.path, subject: saved.subject, senderEmail: saved.senderEmail };

        if (extracted.type === 'image') {
            return contents(
                {
                    type: 'text',
                    text: JSON.stringify({ ...about, type: 'image', mimeType: extracted.mimeType }, null, 2)
                },
                { type: 'image', data: extracted.data, mimeType: extracted.mimeType },
            );
        }
        if (!extracted.text) {
            return json({ ...about, type: extracted.format, text: '', note: extracted.note });
        }
        const truncated = extracted.text.length > max_chars;
        return json({
            ...about,
            type: extracted.format,
            chars: extracted.text.length,
            truncated,
            text: truncated ? extracted.text.slice(0, max_chars) : extracted.text,
        });
    })));
}
