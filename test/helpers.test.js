const test = require('node:test');
const assert = require('node:assert/strict');

const {text, json, safe} = require('../dist/helpers.js');

test('text() wraps a string as one text content block', () => {
    assert.deepEqual(text('hello'), {content: [{type: 'text', text: 'hello'}]});
});

test('json() pretty-prints the value inside a text block', () => {
    const result = json({a: 1});
    assert.equal(result.content.length, 1);
    assert.equal(result.content[0].text, JSON.stringify({a: 1}, null, 2));
});

test('safe() passes through a successful result unchanged', async () => {
    const handler = safe(async (n) => text(`got ${n}`));
    assert.deepEqual(await handler(5), text('got 5'));
});

test('safe() turns a thrown error into a text result instead of rejecting', async () => {
    const handler = safe(async () => {
        throw new Error('boom');
    });
    const result = await handler();
    assert.deepEqual(result, text('Error: boom'));
});
