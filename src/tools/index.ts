// Generic Outlook MCP tools: what @dawnlit/outlook-bridge can do, as tools any
// MCP server can register. Nothing here knows any application's domain — no
// email templates beyond what a mailbox itself stores, no rules about what
// counts as a given kind of email. A host layers those on by registering its
// own tools beside these.
import * as outlook from '@dawnlit/outlook-bridge';
import type { McpServer } from '@modelcontextprotocol/server';
import { registerAttachmentTools } from './attachments';
import { registerCleanupTools } from './cleanup';
import { registerComposeTools } from './compose';
import { registerDraftTools } from './drafts';
import { registerInboxTools } from './inbox';
import {
    OUTLOOK_TOOL_NAMES,
    type OutlookToolName,
    type OutlookToolsOptions,
    type RegisteredOutlookTools,
    type ToolContext,
} from './context';

/**
 * Register the Outlook tools on `server` — every one by default; see
 * `OutlookToolsOptions` to register a read-only set, leave tools out, or point
 * them at a configured bridge. Returns the registered tools by name.
 */
export function registerOutlookTools(server: McpServer, options: OutlookToolsOptions = {}): RegisteredOutlookTools {
    const exclude = new Set<string>(options.exclude ?? []);
    const unknown = [...exclude].filter(name => !(OUTLOOK_TOOL_NAMES as readonly string[]).includes(name));
    if (unknown.length > 0) {
        throw new TypeError(`Unknown Outlook tool(s) to exclude: ${unknown.join(', ')}. Known tools: ${OUTLOOK_TOOL_NAMES.join(', ')}.`);
    }

    const registered: RegisteredOutlookTools = {};
    const context: ToolContext = {
        server,
        bridge: options.bridge ?? outlook,
        readOnly: options.readOnly ?? false,
        add(name: OutlookToolName, access, register) {
            if (exclude.has(name) || (access === 'write' && context.readOnly)) return;
            registered[name] = register();
        },
    };

    registerInboxTools(context);
    registerComposeTools(context);
    registerAttachmentTools(context);
    registerDraftTools(context);
    registerCleanupTools(context);
    return registered;
}
