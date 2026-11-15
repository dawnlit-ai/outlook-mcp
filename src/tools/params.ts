// The arguments that name one email, described once for every tool that takes
// them: a model reads the same instruction wherever it meets the pair.
import { z } from 'zod';

export const ENTRY_ID = z.string().describe('The email\'s entryId, from the listing row that found it');
export const STORE_ID = z.string().describe('The storeId from the SAME listing row — it lets the email resolve in any mailbox, not just the default one').optional();
