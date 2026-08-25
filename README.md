# @dawnlit/outlook-mcp

MCP tools for [`@dawnlit/outlook-bridge`](https://github.com/dawnlit-ai/outlook-bridge) — register a full set of Outlook
read/send/reply/draft/template/cleanup tools on any [MCP](https://modelcontextprotocol.io) server with one call. No
app-specific business logic: every tool here is a thin, generic wrapper over the bridge, so it's the same whether the
host is a freight-forwarding assistant, a sales CRM, or anything else that wants an LLM driving a real Outlook client.

## Install

```bash
npm install @dawnlit/outlook-mcp
```

`@dawnlit/outlook-bridge` comes along as a regular dependency. Also requires `@modelcontextprotocol/sdk` and `zod`
(peer dependencies — bring whatever versions your server already uses; `zod` needs `^3.25` or `^4.0`, matching what the
SDK itself accepts).

To track unreleased changes instead of the last published version, install straight off GitHub `master`:

```bash
npm install github:dawnlit-ai/outlook-mcp#master
```

`prepare` runs `npm run build` automatically on install, so `dist/` is built from that checked-out source with no
separate step. npm resolves the branch to a commit and pins it in `package-lock.json`, so a later plain `npm install`
won't pick up new commits on its own — re-run the command above whenever you want to advance to the current tip.

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

`extractAttachmentText(path)` is also exported directly — the PDF/Excel/plain-text extraction and image passthrough
behind `read_outlook_attachment` — for a host that wants the same handling without going through the tool call.

## Development

```bash
npm run build   # tsc → dist/
npm test        # build, then node:test over the pure-logic parts
```

## License

Apache-2.0 © Dawnlit
