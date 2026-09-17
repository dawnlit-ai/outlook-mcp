// Reading an attachment already saved to disk: text out of PDF, Excel (.xlsx)
// and text files, and the raw bytes of an image, for a client to hand to a
// model as an image rather than have anything guess at its contents. Any other
// format is reported as unreadable, never guessed at.
import fs from 'fs';
import path from 'path';
import { readPdfText } from './pdf';
import { readXlsxText } from './xlsx';

/**
 * The image formats an MCP image block can carry to the models that read them.
 * Other images an email carries — a TIFF fax, a HEIC phone photo — are reported
 * rather than sent: a block with their real mime type travels fine and is
 * rejected at the far end, costing the whole request instead of this one read.
 */
const IMAGE_MIME_TYPES: Readonly<Record<string, string>> = {
    '.png': 'image/png',
    '.jpg': 'image/jpeg',
    '.jpeg': 'image/jpeg',
    '.gif': 'image/gif',
    '.webp': 'image/webp',
};
const UNSUPPORTED_IMAGE_EXTENSIONS = ['.bmp', '.tif', '.tiff', '.heic', '.heif'];

/** An image block holds 10 MB of base64, which is this many bytes on disk. */
export const MAX_IMAGE_BYTES = Math.floor((10 * 1024 * 1024 * 3) / 4);

/** Plain-text formats, read as UTF-8. */
const TEXT_EXTENSIONS = ['.txt', '.csv', '.tsv', '.eml', '.json', '.md', '.html', '.htm', '.xml', '.log'];

/** How much of a text file is read; the tools cap what they return well below this. */
const MAX_TEXT_BYTES = 10 * 1024 * 1024;

export interface ExtractedText {
    type: 'text';
    /** The file extension without its dot, or 'unknown'. */
    format: string;
    /** The extracted text; '' when nothing could be extracted. */
    text: string;
    /** Why `text` is empty, when it is. */
    note?: string;
}

export interface ExtractedImage {
    type: 'image';
    /** The file's bytes, base64-encoded, for an image content block. */
    data: string;
    mimeType: string;
}

export type ExtractedAttachment = ExtractedText | ExtractedImage;

function unreadable(format: string, note: string): ExtractedText {
    return { type: 'text', format, text: '', note };
}

function readTextFile(filePath: string): string {
    const size = fs.statSync(filePath).size;
    if (size <= MAX_TEXT_BYTES) return fs.readFileSync(filePath, 'utf8');
    const buffer = Buffer.alloc(MAX_TEXT_BYTES);
    const fd = fs.openSync(filePath, 'r');
    try {
        fs.readSync(fd, buffer, 0, MAX_TEXT_BYTES, 0);
    } finally {
        fs.closeSync(fd);
    }
    return buffer.toString('utf8');
}

/** The readable content of the file at `filePath`: its text, or its image. */
export async function extractAttachmentContent(filePath: string): Promise<ExtractedAttachment> {
    const extension = path.extname(filePath).toLowerCase();
    const format = extension.slice(1) || 'unknown';

    const mimeType = IMAGE_MIME_TYPES[extension];
    if (mimeType) {
        if (fs.statSync(filePath).size > MAX_IMAGE_BYTES) {
            return unreadable(format, `The image is larger than the ${(MAX_IMAGE_BYTES / 1024 / 1024).toFixed(1)} MB an image block holds — open the file at its path.`);
        }
        return { type: 'image', data: fs.readFileSync(filePath).toString('base64'), mimeType };
    }
    if (UNSUPPORTED_IMAGE_EXTENSIONS.includes(extension)) {
        return unreadable(format, `Image blocks carry JPEG, PNG, GIF and WebP only — convert this ${format.toUpperCase()}, or open the file at its path.`);
    }
    if (extension === '.pdf') {
        const text = await readPdfText(filePath);
        return text.trim()
            ? { type: 'text', format, text }
            : unreadable(format, 'The PDF has no text layer — it is probably a scan. Open the file at its path to view it.');
    }
    if (extension === '.xlsx') return { type: 'text', format, text: await readXlsxText(filePath) };
    if (TEXT_EXTENSIONS.includes(extension)) return { type: 'text', format, text: readTextFile(filePath) };
    return unreadable(format, `No text extractor for '${extension || 'files without an extension'}' — open the file at its path.`);
}
