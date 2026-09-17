# @dawnlit/outlook-mcp

[MCP](https://modelcontextprotocol.io) tools for [`@dawnlit/outlook-bridge`](https://github.com/dawnlit-ai/outlook-bridge):
register Outlook reading, searching, replying, drafting, template, attachment and cleanup tools on any MCP server with
one call. Every tool is a thin, generic wrapper over the bridge — no application's business rules — so it serves a
logistics assistant, a sales CRM or a personal agent equally.

Works wherever the bridge does: Outlook on Windows, and legacy Outlook for Mac.

## Install

```bash
npm install @dawnlit/outlook-mcp @modelcontextprotocol/server zod
```

`@dawnlit/outlook-bridge` comes along as a dependency. `@modelcontextprotocol/server` (v2) and `zod` (v4) are peer
dependencies, so the tools register on the same SDK your server already uses.

## Usage

```ts
import { McpServer } from '@modelcontextprotocol/server';
import { StdioServerTransport } from '@modelcontextprotocol/server/stdio';
import { registerOutlookTools } from '@dawnlit/outlook-mcp';

const server = new McpServer({ name: 'my-server', version: '1.0.0' });

registerOutlookTools(server);
// ...register your own tools alongside...

await server.connect(new StdioServerTransport());
```

### Options

```ts
import { createOutlookBridge } from '@dawnlit/outlook-bridge';

const tools = registerOutlookTools(server, {
    // Only tools that leave the mailbox unchanged.
    readOnly: true,
    // Leave specific tools out (an unknown name throws).
    exclude: ['read_outlook_attachment'],
    // Give the tools their own timeouts, temp directory or cancellation — or a fake in tests.
    bridge: createOutlookBridge({ timeoutMs: 120_000 }),
});

tools.list_outlook_inbox?.disable(); // each registered tool is the SDK's handle
```

## Tools

| Tool                         | Access       | What it does                                                                                                         |
|------------------------------|--------------|----------------------------------------------------------------------------------------------------------------------|
| `get_outlook_accounts`       | read         | The mailboxes the Outlook profile can reach.                                                                         |
| `list_outlook_inbox`         | read         | Recent mail from the Inbox root or one folder (including Sent Items, Drafts, …), with previews and attachment names. |
| `read_outlook_email_body`    | read         | One email's plain-text body, split from the quoted thread.                                                           |
| `list_outlook_inbox_folders` | read         | The folders under the Inbox, with item counts.                                                                       |
| `list_outlook_signatures`    | read         | The Outlook signature names on this machine.                                                                         |
| `outlook_templates`          | read / write | Template emails in a mailbox folder: `list`, and `save` a new one.                                                   |
| `save_outlook_attachment`    | read         | Save one attachment to a private temp directory; returns its path.                                                   |
| `read_outlook_attachment`    | read         | An attachment's text (PDF, .xlsx, text files) or its image, inline.                                                  |
| `outlook_drafts`             | read / write | The account's drafts: `list`, and `delete` or `send` chosen ones.                                                    |
| `move_outlook_emails`        | write        | File emails into a folder — including Deleted Items, recoverably.                                                    |
| `reply_outlook_email`        | write        | A threaded reply, from HTML or a template (section, placeholders, signature).                                        |
| `send_outlook_email`         | write        | A new email: to/cc/bcc, attachments, sent or staged as a draft.                                                      |
| `send_all_drafts`            | write        | Send every draft of an account.                                                                                      |
| `delete_outlook_emails`      | write        | Delete mail by id (dry-run first; received and sent mail refused by default), or purge Deleted Items.                |

With `readOnly`, the write tools aren't registered and `outlook_templates` / `outlook_drafts` offer only `list`.

Every tool declares MCP annotations (`readOnlyHint`, `destructiveHint`, `idempotentHint`, `openWorldHint`) so a host can
decide what needs confirmation, and a title for display. The descriptions tell the model the safety rules that matter —
preview before sending, dry-run before deleting, confirm before bulk actions.

### Errors

A failed call returns a tool result with `isError: true`, its text led by the bridge's error code —
`[ACCOUNT_NOT_FOUND] …`, `[NOT_FOUND] Folder 'Invoices' not found. …`, `[INVALID_REQUEST] …` — so a model can act on the
kind of failure rather than parse its wording.

## Also exported

- `extractAttachmentContent(path)` — the file reading behind `read_outlook_attachment`: text from PDF, .xlsx and text
  formats, or an image's bytes for an image block.
- `readPdfText(path)`, `readXlsxText(path)` — the two extractors on their own. `readPdfText` is also available as
  `@dawnlit/outlook-mcp/pdf`, which loads nothing else.
- `OUTLOOK_TOOL_NAMES` and the `OutlookToolName`, `OutlookToolsOptions` and `RegisteredOutlookTools` types.

## What's deliberately not here

Anything encoding one application's conventions rather than Outlook's own — what counts as a particular kind of email,
an app's own stored templates, rules for when to block a sender. Build those as your own tools beside these, calling
`@dawnlit/outlook-bridge` directly for anything this package doesn't cover.

## Development

```bash
npm run build   # tsc → dist/
npm test        # build, then node:test — no Outlook needed
```

The tests register the tools against a recording server with a fake bridge (what each tool declares and hands the
bridge), and against a real `McpServer` over the SDK's in-memory transport (what a client actually receives).

## License

Apache-2.0 © Dawnlit
