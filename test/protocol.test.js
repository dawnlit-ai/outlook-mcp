// The tools on a real McpServer, over the SDK's in-memory transport, spoken to
// in raw JSON-RPC: what an MCP client actually receives. Guards the parts a
// recording fake can't — that the SDK accepts every definition, turns each zod
// schema into JSON Schema, and delivers error results intact.
const test = require('node:test');
const assert = require('node:assert/strict');

const {McpServer, InMemoryTransport, LATEST_PROTOCOL_VERSION} = require('@modelcontextprotocol/server');
const {AccountNotFoundError} = require('@dawnlit/outlook-bridge');
const {registerOutlookTools} = require('../dist/index.js');

/** A connected server and a function to send it requests. */
async function connect(options) {
    const server = new McpServer({name: 'test', version: '1.0.0'});
    registerOutlookTools(server, options);
    const [client, serverSide] = InMemoryTransport.createLinkedPair();
    const pending = new Map();
    client.onmessage = message => {
        const resolve = pending.get(message.id);
        if (resolve) {
            pending.delete(message.id);
            resolve(message);
        }
    };
    await server.connect(serverSide);
    await client.start();
    let nextId = 1;
    const request = (method, params) => new Promise(resolve => {
        const id = nextId++;
        pending.set(id, resolve);
        client.send({jsonrpc: '2.0', id, method, params});
    });
    const init = await request('initialize', {
        protocolVersion: LATEST_PROTOCOL_VERSION,
        capabilities: {},
        clientInfo: {name: 'test', version: '1.0.0'},
    });
    assert.equal(init.error, undefined, JSON.stringify(init.error));
    await client.send({jsonrpc: '2.0', method: 'notifications/initialized'});
    return {request, close: () => server.close()};
}

const bridge = {
    readInboxEmails: async (account, options) => [{entryId: 'E1', account, options}],
    getOutlookAccounts: async () => {
        throw new AccountNotFoundError('nobody@example.com');
    },
};

test('tools/list describes every tool with a title, annotations and an object schema', async () => {
    const {request, close} = await connect({bridge});
    try {
        const {result} = await request('tools/list', {});
        assert.equal(result.tools.length, 14);
        for (const tool of result.tools) {
            assert.ok(tool.title, tool.name);
            assert.equal(typeof tool.annotations.readOnlyHint, 'boolean', tool.name);
            assert.equal(tool.inputSchema.type, 'object', tool.name);
        }
    } finally {
        await close();
    }
});

test('in read-only mode the listed schemas only offer read actions', async () => {
    const {request, close} = await connect({bridge, readOnly: true});
    try {
        const {result} = await request('tools/list', {});
        const drafts = result.tools.find(tool => tool.name === 'outlook_drafts');
        assert.deepEqual(drafts.inputSchema.properties.action.enum, ['list']);
        assert.ok(!result.tools.some(tool => tool.name === 'send_outlook_email'));
    } finally {
        await close();
    }
});

test('tools/call applies schema defaults and returns the bridge result', async () => {
    const {request, close} = await connect({bridge});
    try {
        const {result} = await request('tools/call', {
            name: 'list_outlook_inbox',
            arguments: {email_account: 'a@example.com', days_back: 3}
        });
        assert.notEqual(result.isError, true);
        const [row] = JSON.parse(result.content[0].text);
        assert.deepEqual(row.options, {daysBack: 3, limit: 50, previewChars: 600});
    } finally {
        await close();
    }
});

test('a failed call reaches the client as an error result with its code', async () => {
    const {request, close} = await connect({bridge});
    try {
        const {result} = await request('tools/call', {name: 'get_outlook_accounts', arguments: {}});
        assert.equal(result.isError, true);
        assert.match(result.content[0].text, /^\[ACCOUNT_NOT_FOUND\] /);
    } finally {
        await close();
    }
});
