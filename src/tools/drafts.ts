// What's staged in the Drafts folder: listing it, retracting or sending chosen
// drafts, and sending everything at once.
import { z } from 'zod';
import { InvalidRequestError } from '@dawnlit/outlook-bridge';
import { json, toolHandler } from '../results';
import type { ToolContext } from './context';

export function registerDraftTools({ server, bridge, readOnly, add }: ToolContext): void {
    const actions = readOnly ? (['list'] as const) : (['list', 'delete', 'send'] as const);
    add('outlook_drafts', 'read', () => server.registerTool('outlook_drafts', {
        title: 'Outlook drafts',
        description:
            "The mail drafts staged for an account — what is actually sitting in its Drafts folder. " +
            "'list' is the only way to see what a batch of staged mail really contains, and answers what a staging run can't tell on its own: whether an earlier batch is still pending, and whether re-drafting left TWO drafts for one recipient (re-drafting adds a draft, it never replaces one). " +
            'Each row gives entryId, subject, To, a short bodyPreview and the folder, newest first; the preview is what tells templated variants apart. Full bodies are never returned.' +
            (readOnly ? '' :
                " 'delete' removes drafts by entryId into Deleted Items (recoverable there) — use it to retract superseded drafts. " +
                "'send' sends only the drafts named by `entry_ids`: prefer it to send_all_drafts whenever the user reviews the list and picks what goes. " +
                "Both verify every id is a draft belonging to THIS account first; an id pointing at ordinary mail, or at another account's draft, is refused and reported. " +
                "⚠️ Deleting or sending someone's mail is visible and hard to undo — list first, tell the user exactly which drafts would be affected, and get their confirmation."),
        inputSchema: z.object({
            action: z.enum(actions).describe(readOnly
                ? "'list' = the account's drafts."
                : "'list' = the account's drafts (read-only). 'delete' = remove drafts by id; needs `entry_ids`. 'send' = send only the named drafts; needs `entry_ids`."),
            email_account: z.string().describe('Outlook email account address whose drafts to act on'),
            entry_ids: z.array(z.string()).min(1).max(100).describe("'delete'/'send' only: entryIds from a preceding 'list' (max 100 per call)").optional(),
            limit: z.number().int().min(1).max(500).describe("'list' only: maximum drafts, newest first (default 100); `count` reports the true total and `truncated` whether it bit").default(100),
            preview_chars: z.number().int().min(0).max(2000).describe("'list' only: characters of body preview per draft (default 300) — raise it when drafts share a long greeting").default(300),
        }),
        annotations: readOnly
            ? { readOnlyHint: true, idempotentHint: true, openWorldHint: false }
            : { readOnlyHint: false, destructiveHint: true, idempotentHint: false, openWorldHint: true },
    }, toolHandler(async ({ action, email_account, entry_ids, limit, preview_chars }) => {
        if (action === 'list') {
            return json(await bridge.listOutlookDrafts(email_account, { limit, previewChars: preview_chars }));
        }
        if (!entry_ids || entry_ids.length === 0) {
            throw new InvalidRequestError(`action '${action}' needs \`entry_ids\` — list the drafts first and pass the ids you mean.`);
        }
        if (action === 'send') {
            return json(await bridge.sendDrafts(email_account, entry_ids));
        }
        return json({
            ...await bridge.deleteOutlookDrafts(email_account, entry_ids),
            note: "Deleted drafts are in the Deleted Items folder, recoverable from there.",
        });
    })));

    add('send_all_drafts', 'write', () => server.registerTool('send_all_drafts', {
        title: 'Send all Outlook drafts',
        description:
            "Send every mail draft bound to the account — each goes out IMMEDIATELY. This is account-wide and batch-blind: it sends whatever it finds, not just a batch you staged. " +
            "Before calling: list the drafts with outlook_drafts (action 'list') and reconcile them against what you staged — more drafts than you expect means unrelated or superseded ones would go too — then tell the user it will send all of them and get their confirmation. " +
            "Prefer outlook_drafts (action 'send') whenever the user wants to choose what goes. Returns how many were sent and any that failed.",
        inputSchema: z.object({
            email_account: z.string().describe('Outlook email account address whose drafts to send'),
        }),
        annotations: { readOnlyHint: false, destructiveHint: true, idempotentHint: false, openWorldHint: true },
    }, toolHandler(async ({ email_account }) => json(await bridge.sendAllDrafts(email_account)))));
}
