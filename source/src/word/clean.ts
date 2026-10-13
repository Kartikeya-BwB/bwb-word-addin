// "Clean up the selected text": the safe house-style wording fixes, applied in place so formatting survives.

import { quoteIssues, scan } from '../rules/writing.ts'
import { replaceSpans } from './fixes.ts'
import type { Span } from './fixes.ts'
import { runWord } from './run.ts'

const HEADING = /^Heading[1-9]$/

export async function cleanSelectionWording(): Promise<string> {
  const result = await runWord(async (context) => {
    const paragraphs = context.document.getSelection().paragraphs
    paragraphs.load('items/text,items/style,items/styleBuiltIn,items/tableNestingLevel')
    await context.sync()
    let applied = 0
    let skipped = 0
    for (const p of paragraphs.items) {
      const text = p.text
      if (!text.trim() || /^BwB Pull Quote/.test(p.style)) continue
      const ctx = { title: HEADING.test(p.styleBuiltIn) || p.styleBuiltIn === 'Title', compact: p.tableNestingLevel > 0 }
      const spans: Span[] = [...scan(text, ctx), ...quoteIssues(text)]
        .filter((i) => i.fixable && i.replacement !== undefined)
        .map((i) => ({ start: i.start, end: i.end, replacement: i.replacement ?? '' }))
      if (!spans.length) continue
      const r = await replaceSpans(context, p, text, spans)
      applied += r.applied
      skipped += r.skipped
    }
    return { applied, skipped, paragraphs: paragraphs.items.length }
  })
  if (!result.applied && !result.skipped) return 'Nothing to change in the selection.'
  const changed = `${result.applied} ${result.applied === 1 ? 'change' : 'changes'} made`
  return result.skipped ? `${changed}; ${result.skipped} left for you to check.` : `${changed}.`
}
