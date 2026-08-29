const test = require('node:test');
const assert = require('node:assert/strict');
const fs = require('fs');
const os = require('os');
const path = require('path');
const ExcelJS = require('exceljs');

const {extractAttachmentText} = require('../dist/attachmentText.js');

function writeTemp(name, contents) {
    const p = path.join(os.tmpdir(), `outlook-mcp-test-${Date.now()}-${name}`);
    fs.writeFileSync(p, contents, 'utf-8');
    return p;
}

test('extractAttachmentText reads a plain-text attachment verbatim', async () => {
    const p = writeTemp('note.txt', 'line one\nline two');
    try {
        const result = await extractAttachmentText(p);
        assert.equal(result.type, 'txt');
        assert.equal(result.text, 'line one\nline two');
        assert.equal(result.note, undefined);
    } finally {
        fs.unlinkSync(p);
    }
});

test('extractAttachmentText returns an image as base64 with its mime type', async () => {
    const p = writeTemp('photo.png', 'not a real png, just bytes');
    try {
        const result = await extractAttachmentText(p);
        assert.equal(result.type, 'image');
        assert.equal(result.mimeType, 'image/png');
        assert.equal(result.data, Buffer.from('not a real png, just bytes').toString('base64'));
    } finally {
        fs.unlinkSync(p);
    }
});

test('extractAttachmentText reports an unhandled extension rather than guessing', async () => {
    const p = writeTemp('archive.zip', 'PK...');
    try {
        const result = await extractAttachmentText(p);
        assert.equal(result.type, 'zip');
        assert.equal(result.text, '');
        assert.match(result.note, /no text extractor/i);
    } finally {
        fs.unlinkSync(p);
    }
});

test('extractAttachmentText keeps a sparse row under its own columns', async () => {
    const p = path.join(os.tmpdir(), `outlook-mcp-test-${Date.now()}-quote.xlsx`);
    const wb = new ExcelJS.Workbook();
    const ws = wb.addWorksheet('Quote');
    ws.getCell('A1').value = 'Charge';
    ws.getCell('B1').value = 'Legal';
    ws.getCell('C1').value = 'OW';
    ws.getCell('D1').value = 'Notes';
    ws.getCell('A2').value = 'Line Haul';
    ws.getCell('D2').value = 'per diem excl';
    await wb.xlsx.writeFile(p);
    try {
        const result = await extractAttachmentText(p);
        assert.equal(result.type, 'xlsx');
        assert.equal(
            result.text,
            '# Quote\nCharge\tLegal\tOW\tNotes\nLine Haul\t\t\tper diem excl',
        );
    } finally {
        fs.unlinkSync(p);
    }
});
