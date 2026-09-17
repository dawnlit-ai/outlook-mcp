// The tools as a host sees them: which get registered, what each declares, and
// what each hands the bridge. Runs without Outlook — the server is a recorder
// and the bridge a fake.
const test = require('node:test');
const assert = require('node:assert/strict');

const {registerOutlookTools, OUTLOOK_TOOL_NAMES} = require('../dist/index.js');
const {NotFoundError, AccountNotFoundError} = require('@dawnlit/outlook-bridge');

/** Records registerTool calls the way McpServer receives them. */
function recordingServer() {
    const tools = new Map();
    return {
        tools,
        registerTool(name, config, handler) {
            assert.ok(!tools.has(name), `${name} registered twice`);
            tools.set(name, {config, handler});
            return {name};
        },
    };
}

/** A bridge whose every method records its arguments and answers `answers[name]`. */
function fakeBridge(answers = {}) {
    const calls = [];
    return new Proxy({calls}, {
        get(target, name) {
            if (name === 'calls') return calls;
            return async (...args) => {
                calls.push({name, args});
                const answer = answers[name];
                return typeof answer === 'function' ? answer(...args) : answer;
            };
        },
    });
}

/** Parse arguments through a tool's own schema (applying its defaults) and run it. */
async function call(server, name, args = {}) {
    const tool = server.tools.get(name);
    assert.ok(tool, `${name} is registered`);
    return tool.handler(tool.config.inputSchema.parse(args));
}

const READ_ONLY_TOOLS = [
    'get_outlook_accounts', 'list_outlook_inbox', 'read_outlook_email_body', 'list_outlook_inbox_folders',
    'list_outlook_signatures', 'outlook_templates', 'save_outlook_attachment', 'read_outlook_attachment', 'outlook_drafts',
];

test('every tool registers by default, and the result names them', () => {
    const server = recordingServer();
    const registered = registerOutlookTools(server, {bridge: fakeBridge()});
    assert.deepEqual([...server.tools.keys()].sort(), [...OUTLOOK_TOOL_NAMES].sort());
    assert.deepEqual(Object.keys(registered).sort(), [...OUTLOOK_TOOL_NAMES].sort());
});

test('every tool declares a title, a description, an input schema and annotations', () => {
    const server = recordingServer();
    registerOutlookTools(server, {bridge: fakeBridge()});
    for (const [name, {config}] of server.tools) {
        assert.ok(config.title, `${name} has a title`);
        assert.ok(config.description.length > 40, `${name} has a description`);
        assert.equal(typeof config.inputSchema.parse, 'function', `${name} has a schema`);
        assert.equal(typeof config.annotations.readOnlyHint, 'boolean', `${name} says whether it is read-only`);
        assert.doesNotMatch(config.description, /windows only/i, `${name} works on macOS too`);
    }
});

test('annotations mark exactly the tools that change nothing as read-only', () => {
    const server = recordingServer();
    registerOutlookTools(server, {bridge: fakeBridge()});
    const readOnly = [...server.tools].filter(([, {config}]) => config.annotations.readOnlyHint).map(([name]) => name);
    assert.deepEqual(readOnly.sort(), READ_ONLY_TOOLS.filter(n => n !== 'outlook_templates' && n !== 'outlook_drafts').sort());
    for (const name of ['delete_outlook_emails', 'send_all_drafts', 'outlook_drafts']) {
        assert.equal(server.tools.get(name).config.annotations.destructiveHint, true, name);
    }
});

test('readOnly registers only tools that leave the mailbox unchanged, and only their read actions', async () => {
    const server = recordingServer();
    registerOutlookTools(server, {bridge: fakeBridge(), readOnly: true});
    assert.deepEqual([...server.tools.keys()].sort(), [...READ_ONLY_TOOLS].sort());
    for (const name of ['outlook_templates', 'outlook_drafts']) {
        const {config} = server.tools.get(name);
        assert.equal(config.annotations.readOnlyHint, true, name);
        assert.throws(() => config.inputSchema.parse({
            action: name === 'outlook_drafts' ? 'send' : 'save',
            email_account: 'a@b.com'
        }), name);
        assert.doesNotMatch(config.description, /'save' adds|'send' sends/);
    }
});

test('exclude leaves tools out, and an unknown name throws instead of registering everything', () => {
    const server = recordingServer();
    registerOutlookTools(server, {bridge: fakeBridge(), exclude: ['send_all_drafts', 'delete_outlook_emails']});
    assert.ok(!server.tools.has('send_all_drafts'));
    assert.ok(!server.tools.has('delete_outlook_emails'));
    assert.ok(server.tools.has('send_outlook_email'));
    assert.throws(() => registerOutlookTools(recordingServer(), {exclude: ['send_all_draft']}), /Unknown Outlook tool/);
});

test('list_outlook_inbox maps its arguments onto the bridge, defaults included', async () => {
    const server = recordingServer();
    const bridge = fakeBridge({readInboxEmails: []});
    registerOutlookTools(server, {bridge});
    await call(server, 'list_outlook_inbox', {email_account: 'a@b.com', folder: 'Sent Items'});
    assert.deepEqual(bridge.calls.at(-1), {
        name: 'readInboxEmails',
        args: ['a@b.com', {daysBack: 60, limit: 50, folder: 'Sent Items', previewChars: 600}],
    });
});

test('send_outlook_email merges both attachment parameters and passes bcc', async () => {
    const server = recordingServer();
    const bridge = fakeBridge();
    registerOutlookTools(server, {bridge});
    const result = await call(server, 'send_outlook_email', {
        email_account: 'a@b.com', to: 'x@y.com', bcc: 'z@y.com', subject: 's', html_body: 'b',
        attachment_path: '/tmp/one.pdf', attachment_paths: ['/tmp/two.pdf'], open_draft_window: false,
    });
    const [params] = bridge.calls.at(-1).args;
    assert.deepEqual(params.attachments, ['/tmp/one.pdf', '/tmp/two.pdf']);
    assert.equal(params.bcc, 'z@y.com');
    assert.equal(result.content[0].text, 'Draft saved to the Outlook Drafts folder.');
});

test('a bridge failure comes back as an error result carrying its code', async () => {
    const server = recordingServer();
    registerOutlookTools(server, {
        bridge: fakeBridge({
            readInboxEmails: () => {
                throw new AccountNotFoundError('nobody@x.com');
            }
        }),
    });
    const result = await call(server, 'list_outlook_inbox', {email_account: 'nobody@x.com'});
    assert.equal(result.isError, true);
    assert.match(result.content[0].text, /^\[ACCOUNT_NOT_FOUND\] /);
});

test('an action missing its ids fails as INVALID_REQUEST without reaching the bridge', async () => {
    const server = recordingServer();
    const bridge = fakeBridge();
    registerOutlookTools(server, {bridge});
    for (const [name, action] of [['outlook_drafts', 'delete'], ['outlook_drafts', 'send'], ['delete_outlook_emails', 'delete']]) {
        const result = await call(server, name, {action, email_account: 'a@b.com'});
        assert.equal(result.isError, true, `${name} ${action}`);
        assert.match(result.content[0].text, /^\[INVALID_REQUEST\] .*entry_ids/);
    }
    assert.equal(bridge.calls.length, 0);
});

test('reply_outlook_email needs a body or a template', async () => {
    const server = recordingServer();
    registerOutlookTools(server, {bridge: fakeBridge()});
    const result = await call(server, 'reply_outlook_email', {email_account: 'a@b.com', entry_id: 'E1'});
    assert.equal(result.isError, true);
});

test('read_outlook_attachment returns an image as an image block beside its metadata', async () => {
    const fs = require('node:fs');
    const os = require('node:os');
    const path = require('node:path');
    const file = path.join(fs.mkdtempSync(path.join(os.tmpdir(), 'outlook-mcp-test-')), 'scan.png');
    fs.writeFileSync(file, 'png bytes');
    const server = recordingServer();
    const bridge = fakeBridge({
        saveEmailAttachment: {
            path: file,
            fileName: 'scan.png',
            subject: 'Scan',
            senderName: '',
            senderEmail: 's@x.com',
            receivedTime: ''
        },
    });
    registerOutlookTools(server, {bridge});
    const result = await call(server, 'read_outlook_attachment', {
        entry_id: 'E1',
        store_id: 'S1',
        file_name: 'scan.png'
    });
    assert.deepEqual(bridge.calls.at(-1).args, [{entryId: 'E1', storeId: 'S1'}, 'scan.png']);
    assert.equal(result.content.length, 2);
    assert.equal(JSON.parse(result.content[0].text).senderEmail, 's@x.com');
    assert.deepEqual(result.content[1], {
        type: 'image',
        data: Buffer.from('png bytes').toString('base64'),
        mimeType: 'image/png'
    });
});

test('read_outlook_attachment reports a missing attachment by code', async () => {
    const server = recordingServer();
    registerOutlookTools(server, {
        bridge: fakeBridge({
            saveEmailAttachment: () => {
                throw new NotFoundError('attachment', "Attachment 'x.pdf' not found.");
            }
        }),
    });
    const result = await call(server, 'read_outlook_attachment', {entry_id: 'E1', file_name: 'x.pdf'});
    assert.equal(result.isError, true);
    assert.match(result.content[0].text, /^\[NOT_FOUND\] Attachment/);
});
