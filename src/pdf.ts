// Text extraction from a PDF on disk. pdf-parse holds a worker open behind the
// parser, so every caller has to destroy it — owning that dance here is the
// point of the function, and the reason a third call site can't get it wrong.
import fs from 'fs';
import { PDFParse } from 'pdf-parse';

/**
 * Extract the full text of the PDF at `filePath`, or '' if it has none.
 * pdf-parse closes every page with a `-- 1 of 3 --` marker, so the joined text
 * of a scan is never empty: whether there is anything to read is asked of the
 * pages themselves.
 */
export async function readPdfText(filePath: string): Promise<string> {
    const parser = new PDFParse({ data: fs.readFileSync(filePath) });
    try {
        const parsed = await parser.getText();
        return parsed.pages.some(page => page.text.trim()) ? parsed.text : '';
    } finally {
        await parser.destroy();
    }
}
