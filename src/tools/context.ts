// What every tool module registers through.
import type { McpServer, RegisteredTool } from '@modelcontextprotocol/server';
import type { OutlookBridge } from '@dawnlit/outlook-bridge';

/** Every tool this package can register. */
export const OUTLOOK_TOOL_NAMES = [
    'get_outlook_accounts',
    'list_outlook_inbox',
    'read_outlook_email_body',
    'list_outlook_inbox_folders',
    'move_outlook_emails',
    'reply_outlook_email',
    'send_outlook_email',
    'list_outlook_signatures',
    'outlook_templates',
    'save_outlook_attachment',
    'read_outlook_attachment',
    'outlook_drafts',
    'send_all_drafts',
    'delete_outlook_emails',
] as const;

export type OutlookToolName = typeof OUTLOOK_TOOL_NAMES[number];

export interface OutlookToolsOptions {
    /**
     * The bridge the tools call. Defaults to the package's exported functions,
     * which run under the process-wide settings; pass
     * `createOutlookBridge({ timeoutMs, tempDir, … })` to give the tools their
     * own, or any object implementing `OutlookBridge` — a fake in tests, say.
     */
    bridge?: OutlookBridge;
    /**
     * Register only what leaves the mailbox unchanged: the reading tools, plus
     * the 'list' actions of `outlook_templates` and `outlook_drafts`. Nothing
     * that sends, files, saves into or deletes from the mailbox is registered.
     */
    readOnly?: boolean;
    /** Tools to leave out, by name. An unknown name throws, so a typo can't silently register a tool. */
    exclude?: readonly OutlookToolName[];
}

/** The tools registered, by name — each one the SDK's handle for enabling, disabling or removing it. */
export type RegisteredOutlookTools = Partial<Record<OutlookToolName, RegisteredTool>>;

/** Whether a tool can change the mailbox or send mail. */
export type ToolAccess = 'read' | 'write';

export interface ToolContext {
    readonly server: McpServer;
    readonly bridge: OutlookBridge;
    readonly readOnly: boolean;

    /** Register a tool unless it is excluded, or writes while `readOnly` is set. */
    add(name: OutlookToolName, access: ToolAccess, register: () => RegisteredTool): void;
}
