# @dawnlit/outlook-mcp

MCP tools for [`@dawnlit/outlook-bridge`](https://github.com/dawnlit-ai/outlook-bridge) — register a full set of Outlook
read/send/reply/draft/template/cleanup tools on any [MCP](https://modelcontextprotocol.io) server with one call. No
app-specific business logic: every tool here is a thin, generic wrapper over the bridge, so it's the same whether the
host is a freight-forwarding assistant, a sales CRM, or anything else that wants an LLM driving a real Outlook client.

## Install

Neither this package nor `@dawnlit/outlook-bridge` is on the npm registry yet — install both straight off GitHub
`main`:

```bash
npm install github:dawnlit-ai/outlook-bridge#main github:dawnlit-ai/outlook-mcp#main
```

Also requires `@modelcontextprotocol/sdk` and `zod` (peer dependencies — bring whatever versions your server already
uses; `zod` needs `^3.25` or `^4.0`, matching what the SDK itself accepts).

## Usage

```ts
import { McpServer } from '@modelcontextprotocol/sdk/server/mcp.js';
import { StdioServerTransport } from '@modelcontextprotocol/sdk/server/stdio.js';
import { registerOutlookTools } from '@dawnlit/outlook-mcp';

const server = new McpServer({ name: 'my-server', version: '1.0.0' });

registerOutlookTools(server);
// ...register your own app-specific tools alongside these...

await server.connect(new StdioServerTransport());
```

That registers: `get_outlook_accounts`, `list_outlook_inbox`, `read_outlook_email_body`,
`list_outlook_inbox_folders`, `move_outlook_emails`, `reply_outlook_email`, `list_outlook_signatures`,
`outlook_templates`, `save_outlook_attachment`, `read_outlook_attachment`, `send_outlook_email`, `outlook_drafts`,
`delete_outlook_emails`, and `send_all_drafts`.

## What's deliberately NOT here

Anything that encodes a particular application's conventions rather than Outlook's own — e.g. "what counts as a
pre-alert email," an app's own config-stored email templates (as opposed to templates stored as mailbox items, which
`outlook_templates` already covers), or business rules about when to blacklist a sender. Compose those as your own MCP
tools on top, calling into `@dawnlit/outlook-bridge` directly for anything this package doesn't cover.

`extractAttachmentText(path)` is also exported directly — the PDF/Excel/plain-text extraction behind
`read_outlook_attachment` — for a host that wants the same extraction without going through the tool call.

## Development

```bash
npm run build   # tsc → dist/
npm test        # build, then node:test over the pure-logic parts
```

## License

Apache-2.0 © Dawnlit
