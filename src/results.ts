// Building tool results, and turning a failure into one.
import type { CallToolResult, ImageContent, TextContent } from '@modelcontextprotocol/server';
import { OutlookError } from '@dawnlit/outlook-bridge';

export function text(value: string): CallToolResult {
    return { content: [{ type: 'text', text: value }] };
}

export function json(value: unknown): CallToolResult {
    return text(JSON.stringify(value, null, 2));
}

/** A result from several content blocks — e.g. JSON metadata beside an inline image. */
export function contents(...blocks: (TextContent | ImageContent)[]): CallToolResult {
    return { content: blocks };
}

/**
 * The stable code of a package error. Checked by shape as well as by class, so
 * an error from a second installed copy of the bridge still reports its code.
 */
function errorCode(error: unknown): string | undefined {
    if (error instanceof OutlookError) return error.code;
    const code = (error as { code?: unknown } | null)?.code;
    return typeof code === 'string' && /^[A-Z_]+$/.test(code) ? code : undefined;
}

/**
 * A failure as a tool result the client can recognize: `isError` set, and the
 * error's code leading the message (`[ACCOUNT_NOT_FOUND] …`) so a model can act
 * on the kind of failure rather than parse its wording.
 */
export function errorResult(error: unknown): CallToolResult {
    const message = error instanceof Error ? error.message : String(error);
    const code = errorCode(error);
    return { isError: true, content: [{ type: 'text', text: code ? `[${code}] ${message}` : message }] };
}

/** Wrap a tool handler so every failure comes back as an error result. */
export function toolHandler<Args>(fn: (args: Args) => Promise<CallToolResult>): (args: Args) => Promise<CallToolResult> {
    return async (args: Args) => {
        try {
            return await fn(args);
        } catch (error) {
            return errorResult(error);
        }
    };
}
