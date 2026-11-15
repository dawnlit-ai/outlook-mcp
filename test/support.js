// What the test files share. Not a test itself: the runner takes `*.test.js`.
const fs = require('node:fs');
const os = require('node:os');
const path = require('node:path');

/** Every directory `tempPath` made, removed when the test process exits. */
const scratchDirs = [];
process.on('exit', () => {
    for (const dir of scratchDirs) fs.rmSync(dir, {recursive: true, force: true});
});

/** A path for `name` in a fresh directory, which is removed when the test process exits. */
function tempPath(name) {
    const dir = fs.mkdtempSync(path.join(os.tmpdir(), 'outlook-mcp-test-'));
    scratchDirs.push(dir);
    return path.join(dir, name);
}

/** A file holding `contents`, in a directory of its own. */
function writeTemp(name, contents) {
    const file = tempPath(name);
    fs.writeFileSync(file, contents);
    return file;
}

/**
 * A minimal PDF with one page per entry of `pageTexts`. An empty entry makes a
 * page with nothing to read on it, as a scanned page has.
 */
function pdfBytes(pageTexts) {
    const fontNumber = pageTexts.length * 2 + 3;
    const objects = ['<< /Type /Catalog /Pages 2 0 R >>', ''];
    const kids = [];
    for (const text of pageTexts) {
        const stream = text ? `BT /F1 18 Tf 20 100 Td (${text}) Tj ET` : '';
        const pageNumber = objects.length + 1;
        kids.push(`${pageNumber} 0 R`);
        objects.push(`<< /Type /Page /Parent 2 0 R /MediaBox [0 0 300 144] /Contents ${pageNumber + 1} 0 R /Resources << /Font << /F1 ${fontNumber} 0 R >> >> >>`);
        objects.push(`<< /Length ${stream.length} >>\nstream\n${stream}\nendstream`);
    }
    objects[1] = `<< /Type /Pages /Kids [${kids.join(' ')}] /Count ${pageTexts.length} >>`;
    objects.push('<< /Type /Font /Subtype /Type1 /BaseFont /Helvetica >>');
    let pdf = '%PDF-1.4\n';
    const offsets = objects.map((object, i) => {
        const offset = pdf.length;
        pdf += `${i + 1} 0 obj\n${object}\nendobj\n`;
        return offset;
    });
    const xref = pdf.length;
    pdf += `xref\n0 ${objects.length + 1}\n0000000000 65535 f \n`
        + offsets.map(offset => `${String(offset).padStart(10, '0')} 00000 n \n`).join('')
        + `trailer\n<< /Size ${objects.length + 1} /Root 1 0 R >>\nstartxref\n${xref}\n%%EOF\n`;
    return Buffer.from(pdf, 'latin1');
}

module.exports = {tempPath, writeTemp, pdfBytes};
