// Refreshing fields: contents, caption numbers (SEQ), cross-references.

import { runWord, supports, UserFacingError } from './run.ts'

/** Updates every field in the body, or those whose code passes the filter. Needs WordApi 1.5. */
export async function updateFields(filter?: (code: string) => boolean): Promise<number> {
  if (!supports('1.5')) {
    throw new UserFacingError('Updating fields needs a newer Word. Press Ctrl+A then F9 to update them by hand.')
  }
  return runWord(async (context) => {
    const fields = context.document.body.fields
    fields.load('items/code')
    await context.sync()
    const targets = fields.items.filter((f) => !filter || filter(f.code))
    targets.forEach((f) => f.updateResult())
    await context.sync()
    return targets.length
  })
}

export const isSeqField = (code: string): boolean => /^\s*SEQ\b/i.test(code)
export const isTocField = (code: string): boolean => /^\s*TOC\b/i.test(code)
