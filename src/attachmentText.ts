// Best-effort attachment reading for an attachment already saved to disk —
// the generic half of "open this Outlook attachment and read it": text
// extraction from PDF, Excel (.xlsx), and plain-text/CSV; raw base64 passthrough
// for images, so the MCP client can hand them to the model as native vision
// input instead of us guessing at their contents. Anything else is reported
// unhandled rather than guessed at.
import fs from 'fs';
import path from 'path';
import { readPdfText } from './pdf';
import { readXlsxText } from './xlsx';

// The four formats Claude's vision accepts. The others an email actually carries —
// a TIFF fax of a B/L, a HEIC phone photo — are listed separately below rather than
// mapped to their real mime type: MCP would transport such a block happily (mimeType
// is a free string) and the API would reject it, costing the whole request instead of
// just this tool.
const IMAGE_MIME_TYPES: Record<string, string> = {
    '.png': 'image/png',
    '.jpg': 'image/jpeg',
    '.jpeg': 'image/jpeg',
    '.gif': 'image/gif',
    '.webp': 'image/webp',
};
const UNREADABLE_IMAGE_EXTS = ['.bmp', '.tif', '.tiff', '.heic'];
/** An image block accepts 10 MB of base64, which is this many bytes on disk. */
const MAX_IMAGE_BYTES = Math.floor((10 * 1024 * 1024 * 3) / 4);
const TEXT_EXTS = ['.txt', '.csv', '.tsv', '.eml', '.json', '.md'];

export interface ExtractedAttachmentText {
    /** File extension without the leading dot, or 'unknown'. */
    type: string;
    text: string;
    /** Set when `text` is empty because nothing could be extracted. */
    note?: string;
}

export interface ExtractedAttachmentImage {
    type: 'image';
    /** Raw bytes, base64-encoded, for a native MCP image content block. */
    data: string;
    mimeType: string;
}

export type ExtractedAttachment = ExtractedAttachmentText | ExtractedAttachmentImage;

export async function extractAttachmentText(savedPath: string): Promise<ExtractedAttachment> {
    const ext = path.extname(savedPath).toLowerCase();

    const imageMimeType = IMAGE_MIME_TYPES[ext];
    if (imageMimeType) {
        if (fs.statSync(savedPath).size > MAX_IMAGE_BYTES) {
            return {
                type: ext.replace('.', ''),
                text: '',
                note: `Image is over the ${(MAX_IMAGE_BYTES / 1024 / 1024).toFixed(1)} MB an image block holds — open the file at the path.`,
            };
        }
        return { type: 'image', data: fs.readFileSync(savedPath).toString('base64'), mimeType: imageMimeType };
    }

    if (UNREADABLE_IMAGE_EXTS.includes(ext)) {
        return {
            type: ext.replace('.', ''),
            text: '',
            note: `Claude reads JPEG, PNG, GIF and WebP only — convert this ${ext.replace('.', '').toUpperCase()}, or open the file at the path.`,
        };
    }

    if (ext === '.pdf') {
        return { type: 'pdf', text: await readPdfText(savedPath) };
    }

    if (ext === '.xlsx') {
        return { type: 'xlsx', text: await readXlsxText(savedPath) };
    }

    if (TEXT_EXTS.includes(ext)) {
        return { type: ext.replace('.', ''), text: fs.readFileSync(savedPath, 'utf-8') };
    }

    return {
        type: ext.replace('.', '') || 'unknown',
        text: '',
        note: `No text extractor for '${ext}' — open the file manually at the path.`,
    };
}
