const test = require('node:test');
const assert = require('node:assert/strict');
const fs = require('fs');
const os = require('os');
const path = require('path');

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

test('extractAttachmentText reports an image as having no extractable text', async () => {
    const p = writeTemp('photo.png', 'not a real png, just bytes');
    try {
        const result = await extractAttachmentText(p);
        assert.equal(result.type, 'image');
        assert.equal(result.text, '');
        assert.match(result.note, /no extractable text/i);
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
