// Text extraction from a PDF on disk. pdf-parse holds a worker open behind the
// parser, so every caller has to destroy it — owning that dance here is the
// point of the function, and the reason a third call site can't get it wrong.
import fs from 'fs';
import { PDFParse } from 'pdf-parse';

/** Extract the full text of the PDF at `filePath`, or '' if it has none. */
export async function readPdfText(filePath: string): Promise<string> {
    const parser = new PDFParse({ data: fs.readFileSync(filePath) });
    try {
        const parsed = await parser.getText();
        return parsed.text || '';
    } finally {
        await parser.destroy();
    }
}
