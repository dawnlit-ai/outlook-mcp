// Best-effort attachment reading for an attachment already saved to disk —
// the generic half of "open this Outlook attachment and read it": text
// extraction from PDF, Excel (.xlsx), and plain-text/CSV; raw base64 passthrough
// for images, so the MCP client can hand them to the model as native vision
// input instead of us guessing at their contents. Anything else is reported
// unhandled rather than guessed at.
import fs from 'fs';
import path from 'path';
import ExcelJS from 'exceljs';
import { readPdfText } from './pdf';

const IMAGE_MIME_TYPES: Record<string, string> = {
    '.png': 'image/png',
    '.jpg': 'image/jpeg',
    '.jpeg': 'image/jpeg',
    '.gif': 'image/gif',
    '.bmp': 'image/bmp',
    '.tif': 'image/tiff',
    '.tiff': 'image/tiff',
    '.webp': 'image/webp',
    '.heic': 'image/heic',
};
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
        return { type: 'image', data: fs.readFileSync(savedPath).toString('base64'), mimeType: imageMimeType };
    }

    if (ext === '.pdf') {
        return { type: 'pdf', text: await readPdfText(savedPath) };
    }

    if (ext === '.xlsx') {
        const wb = new ExcelJS.Workbook();
        await wb.xlsx.readFile(savedPath);
        const lines: string[] = [];
        wb.worksheets.forEach((ws) => {
            lines.push(`# ${ws.name}`);
            ws.eachRow({ includeEmpty: false }, (row) => {
                const vals: string[] = [];
                row.eachCell({ includeEmpty: false }, (cell) => {
                    const v = String(cell.text ?? '').trim();
                    if (v) vals.push(v);
                });
                if (vals.length) lines.push(vals.join('\t'));
            });
        });
        return { type: 'xlsx', text: lines.join('\n') };
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
