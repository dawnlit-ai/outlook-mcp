// Managing what's sitting in Outlook's Drafts folder: listing, retracting
// (delete), sending a chosen subset, and sending everything at once.
import { z } from 'zod';
import type { McpServer } from '@modelcontextprotocol/server';
import { deleteOutlookDrafts, listOutlookDrafts, sendAllDrafts, sendDrafts, } from '@dawnlit/outlook-bridge';
import { json, safe } from '../helpers';

/** Register the drafts list/delete/send tool and the send-all-drafts tool on `server`. */
export function registerDraftTools(server: McpServer): void {
    server.registerTool('outlook_drafts', {
        description:
            "The mail drafts staged for an account — what is actually sitting in Outlook's Drafts folder(s), and the ways to remove or send one. " +
            "'list' is READ-ONLY and is the only way to see what a batch of staged mail really contains: call it BEFORE send_all_drafts, which is account-wide and sends everything it finds. " +
            "It answers the two questions a draft-staging run cannot answer for itself — whether an earlier batch is still pending or was already sent, and whether re-drafting has left TWO drafts for one recipient. " +
            'Each row gives entryId, subject, To, a short bodyPreview and the folder, newest first; the preview is what distinguishes one templated variant from another. Full bodies are never returned — a templated body can run tens of thousands of characters. ' +
            "'delete' removes drafts by entryId. Outlook moves them to Deleted Items, so a mistake is recoverable there, and every id is verified to be a mail draft belonging to THIS account before anything happens — an id pointing at ordinary inbox mail is refused, not deleted. " +
            "Use it to retract superseded drafts: re-drafting a reply does NOT replace the earlier one, it adds a second, and bulk-sending would then send both. " +
            "'send' sends only the drafts named by `entry_ids` — the reviewed subset from a 'list' call — instead of send_all_drafts' account-wide sweep. Prefer this whenever the user wants to see what would go out and pick some of it, rather than flush the whole folder. Same ownership gate as 'delete': an id pointing at ordinary mail, or at another account's draft, is refused and reported rather than sent. " +
            '⚠️ Deleting or sending the user\'s mail is a visible, hard-to-undo action — list first, tell them exactly which drafts would go, and get explicit confirmation. Never delete or send a draft the user has not agreed to. ' +
            'Drafts bound to a different account are invisible to every action. Windows only (Outlook COM).',
        inputSchema: z.object({
            action: z.enum(['list', 'delete', 'send']).describe("'list' = the account's drafts (read-only). 'delete' = remove drafts by id; needs `entry_ids`. 'send' = send only the named drafts; needs `entry_ids`."),
            email_account: z.string().describe('Outlook email account address whose drafts to list, delete, or send'),
            entry_ids: z.array(z.string()).min(1).max(100).describe("'delete'/'send' only: EntryIDs from a preceding 'list' (max 100 per call)").optional(),
            limit: z.number().int().min(1).max(500).describe("'list' only: maximum drafts to return, newest first (default 100); `count` reports the true total and `truncated` says whether it bit").default(100),
            preview_chars: z.number().int().min(0).max(2000).describe("'list' only: characters of plain-text body preview per draft (default 300). Raise it when several drafts share a greeting and you need to read further in to tell them apart.").default(300),
        }),
    }, safe(async ({ action, email_account, entry_ids, limit, preview_chars }) => {
        if (action === 'list') {
            return json(await listOutlookDrafts(email_account, limit, preview_chars));
        }
        if (!entry_ids || entry_ids.length === 0) {
            throw new Error(`action '${action}' needs \`entry_ids\` — list the drafts first and pass the ids you mean.`);
        }
        if (action === 'send') {
            const res = await sendDrafts(email_account, entry_ids);
            return json(res);
        }
        const res = await deleteOutlookDrafts(email_account, entry_ids);
        return json({
            ...res,
            note: "Deleted drafts are in Outlook's Deleted Items folder, not gone — recoverable from there.",
        });
    }));

    server.registerTool('send_all_drafts', {
        description: "Send every mail draft in Outlook's Drafts folder that is set to send from the given account. Each matching draft goes out IMMEDIATELY, so this is a bulk send: before calling, tell the user it will send all queued drafts for that account at once and get explicit confirmation. It is batch-blind — it cannot send only a batch you just staged. Call `outlook_drafts` (action 'list') FIRST and reconcile what's there against what you staged: a count higher than yours means unrelated or superseded drafts would go out too, and superseded ones should be removed with `outlook_drafts` (action 'delete') before sending. Prefer `outlook_drafts` (action 'send') instead whenever the user wants to review the list and pick which ones go — it sends only the entry_ids you give it. Drafts bound to a different account, or with no sending account set, are left untouched. Returns how many were sent and any that failed (by subject). Use get_outlook_accounts first if you're unsure of the exact account address.",
        inputSchema: z.object({
            email_account: z.string().describe('Outlook email account address (sender) whose drafts should be sent'),
        }),
    }, safe(async ({ email_account }) => {
        const res = await sendAllDrafts(email_account);
        return json(res);
    }));
}
