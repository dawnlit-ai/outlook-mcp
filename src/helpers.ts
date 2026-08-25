// Minimal MCP tool-result helpers. Deliberately not imported from anywhere else —
// small enough to own here rather than add a dependency for a handful of functions.

export interface ToolTextContent {
    type: 'text';
    text: string;
}

export interface ToolImageContent {
    type: 'image';
    data: string;
    mimeType: string;
}

export type ToolContent = ToolTextContent | ToolImageContent;

export interface ToolResult {
    // Matches the MCP SDK's CallToolResult shape (which carries an index
    // signature alongside its known fields) so a ToolResult is assignable to
    content: ToolContent[];

    // it without every caller re-declaring the SDK's own type.
    [key: string]: unknown;
}

export function text(value: string): ToolResult {
    return { content: [{ type: 'text', text: value }] };
}

export function json(value: unknown): ToolResult {
    return text(JSON.stringify(value, null, 2));
}

/** A ToolResult from any mix of content blocks — e.g. JSON metadata alongside an inline image. */
export function contents(...blocks: ToolContent[]): ToolResult {
    return { content: blocks };
}

/** Wrap a tool handler so a thrown error becomes a text result instead of crashing the server. */
export function safe<Args extends unknown[]>(
    fn: (...args: Args) => Promise<ToolResult>,
): (...args: Args) => Promise<ToolResult> {
    return async (...args: Args) => {
        try {
            return await fn(...args);
        } catch (e) {
            return text(`Error: ${(e as Error).message}`);
        }
    };
}
