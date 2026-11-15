const test = require('node:test');
const assert = require('node:assert/strict');
const ExcelJS = require('exceljs');

const {extractAttachmentContent} = require('../dist/attachmentContent.js');
const {tempPath, writeTemp, pdfBytes} = require('./support');

test('a plain-text attachment is read verbatim', async () => {
    const result = await extractAttachmentContent(writeTemp('note.txt', 'line one\nline two'));
    assert.deepEqual(result, {type: 'text', format: 'txt', text: 'line one\nline two'});
});

test('an HTML attachment is read as text', async () => {
    const result = await extractAttachmentContent(writeTemp('page.html', '<p>Hi</p>'));
    assert.equal(result.type, 'text');
    assert.equal(result.text, '<p>Hi</p>');
});

test('an image comes back as base64 with its mime type', async () => {
    const result = await extractAttachmentContent(writeTemp('photo.png', 'not a real png, just bytes'));
    assert.equal(result.type, 'image');
    assert.equal(result.mimeType, 'image/png');
    assert.equal(result.data, Buffer.from('not a real png, just bytes').toString('base64'));
});

test('an unhandled extension is reported rather than guessed at', async () => {
    const result = await extractAttachmentContent(writeTemp('archive.zip', 'PK...'));
    assert.equal(result.type, 'text');
    assert.equal(result.format, 'zip');
    assert.equal(result.text, '');
    assert.match(result.note, /no text extractor/i);
});

test('a sparse spreadsheet row keeps each value under its own column', async () => {
    const file = tempPath('sheet.xlsx');
    const workbook = new ExcelJS.Workbook();
    const sheet = workbook.addWorksheet('Prices');
    sheet.getCell('A1').value = 'Item';
    sheet.getCell('B1').value = 'Small';
    sheet.getCell('C1').value = 'Large';
    sheet.getCell('D1').value = 'Notes';
    sheet.getCell('A2').value = 'Widget';
    sheet.getCell('D2').value = 'on request';
    await workbook.xlsx.writeFile(file);
    const result = await extractAttachmentContent(file);
    assert.equal(result.format, 'xlsx');
    assert.equal(result.text, '# Prices\nItem\tSmall\tLarge\tNotes\nWidget\t\t\ton request');
});

test("a PDF's text is read, each page closed by its marker", async () => {
    const result = await extractAttachmentContent(writeTemp('offer.pdf', pdfBytes(['First page', 'Second page'])));
    assert.equal(result.type, 'text');
    assert.equal(result.format, 'pdf');
    assert.match(result.text, /First page\s+-- 1 of 2 --\s+Second page\s+-- 2 of 2 --/);
});

test('a PDF with no text layer is reported as a scan, not returned as its page markers', async () => {
    const result = await extractAttachmentContent(writeTemp('scan.pdf', pdfBytes(['', ''])));
    assert.equal(result.type, 'text');
    assert.equal(result.format, 'pdf');
    assert.equal(result.text, '', 'two bare markers are not text');
    assert.match(result.note, /no text layer/);
});

test('a PDF with one readable page among scanned ones is still read', async () => {
    const result = await extractAttachmentContent(writeTemp('mixed.pdf', pdfBytes(['', 'Signed on page two'])));
    assert.match(result.text, /-- 1 of 2 --\s+Signed on page two/);
    assert.equal(result.note, undefined);
});

test('an image format image blocks cannot carry is reported, not sent', async () => {
    const result = await extractAttachmentContent(writeTemp('scan.tiff', 'II*\u0000 pretend TIFF bytes'));
    assert.equal(result.type, 'text');
    assert.equal(result.format, 'tiff');
    assert.match(result.note, /JPEG, PNG, GIF and WebP only/);
});

test('an image too large for an image block is reported, not read', async () => {
    const result = await extractAttachmentContent(writeTemp('huge.png', Buffer.alloc(8 * 1024 * 1024)));
    assert.equal(result.type, 'text');
    assert.equal(result.data, undefined);
    assert.match(result.note, /open the file at its path/);
});
