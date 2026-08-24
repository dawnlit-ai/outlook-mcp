// Best-effort text extraction for an attachment already saved to disk — the
// generic half of "open this Outlook attachment and read it": PDF, Excel
// (.xlsx), and plain-text/CSV. Images have no extractable text; anything else
// is reported unhandled rather than guessed at.
import fs from 'fs';
import path from 'path';
import { PDFParse } from 'pdf-parse';
import ExcelJS from 'exceljs';

const IMAGE_EXTS = ['.png', '.jpg', '.jpeg', '.gif', '.bmp', '.tif', '.tiff', '.webp', '.heic'];
const TEXT_EXTS = ['.txt', '.csv', '.tsv', '.eml', '.json', '.md'];

export interface ExtractedAttachmentText {
    /** File extension without the leading dot, or 'image'/'unknown'. */
    type: string;
    text: string;
    /** Set when `text` is empty because nothing could be extracted. */
    note?: string;
}

export async function extractAttachmentText(savedPath: string): Promise<ExtractedAttachmentText> {
    const ext = path.extname(savedPath).toLowerCase();

    if (IMAGE_EXTS.includes(ext)) {
        return {
            type: 'image',
            text: '',
            note: 'Image attachment — no extractable text (a scanned or screenshot document needs manual reading/OCR).',
        };
    }

    if (ext === '.pdf') {
        const buf = fs.readFileSync(savedPath);
        const parser = new PDFParse({ data: buf });
        try {
            const parsed = await parser.getText();
            return { type: 'pdf', text: parsed.text || '' };
        } finally {
            await parser.destroy();
        }
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
