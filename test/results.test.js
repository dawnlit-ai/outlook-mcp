const test = require('node:test');
const assert = require('node:assert/strict');

const {text, json, errorResult, toolHandler} = require('../dist/results.js');
const {InvalidRequestError} = require('@dawnlit/outlook-bridge');

test('text() wraps a string as one text block', () => {
    assert.deepEqual(text('hello'), {content: [{type: 'text', text: 'hello'}]});
});

test('json() pretty-prints the value inside a text block', () => {
    assert.equal(json({a: 1}).content[0].text, JSON.stringify({a: 1}, null, 2));
});

test('errorResult sets isError and leads with the error code', () => {
    assert.deepEqual(errorResult(new InvalidRequestError('bad')), {
        isError: true,
        content: [{type: 'text', text: '[INVALID_REQUEST] bad'}],
    });
    assert.deepEqual(errorResult(new Error('plain')), {isError: true, content: [{type: 'text', text: 'plain'}]});
    assert.deepEqual(errorResult('thrown string'), {isError: true, content: [{type: 'text', text: 'thrown string'}]});
});

test('an error from another copy of the bridge still reports its code', () => {
    const foreign = Object.assign(new Error('gone'), {code: 'NOT_FOUND'});
    assert.equal(errorResult(foreign).content[0].text, '[NOT_FOUND] gone');
});

test('toolHandler passes a result through and turns a throw into an error result', async () => {
    assert.deepEqual(await toolHandler(async n => text(`got ${n}`))(5), text('got 5'));
    const failed = await toolHandler(async () => {
        throw new Error('boom');
    })();
    assert.equal(failed.isError, true);
});
