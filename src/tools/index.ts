// Generic Outlook MCP tools: everything @dawnlit/outlook-bridge can do, exposed
// as MCP tools any host server can register. Nothing here knows about any
// particular application's domain — no email templates beyond what a mailbox
// itself stores, no business rules about what counts as a given kind of email.
// A host that wants those layers composes them on top by registering its own
// tools alongside these.
import type { McpServer } from '@modelcontextprotocol/sdk/server/mcp.js';
import { registerAttachmentTools } from './attachments';
import { registerCleanupTools } from './cleanup';
import { registerComposeTools } from './compose';
import { registerDraftTools } from './drafts';
import { registerInboxTools } from './inbox';

/**
 * Register every generic Outlook tool — accounts, inbox reading, folders,
 * replying/sending, drafts, signatures, mailbox-stored templates, attachments,
 * and mailbox cleanup — on `server`. Every tool is backed purely by
 * `@dawnlit/outlook-bridge`; none of it depends on anything the host app owns.
 */
export function registerOutlookTools(server: McpServer): void {
    registerInboxTools(server);
    registerComposeTools(server);
    registerAttachmentTools(server);
    registerDraftTools(server);
    registerCleanupTools(server);
}
