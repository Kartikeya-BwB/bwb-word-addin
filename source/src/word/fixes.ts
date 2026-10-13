// Going to a finding and applying its fix. Locations are paragraph indexes from the last snapshot, so every
// action first checks the paragraph still reads as it did; if not, the author re-runs the check.

import type { FixAction, Location } from '../check/types.ts'
import { addBwbStyles, hasBwbStyles, restyleBuiltInStyles } from './styles.ts'
import { insertBlocks } from './insert.ts'
import { writeDocInfo } from './properties.ts'
import { setSticker } from './stickers.ts'
import { runWord, supports, UserFacingError } from './run.ts'

export class StaleDocumentError extends UserFacingError {
  constructor() {
    super('The document has changed since the check ran. Run the check again.')
  }
}

const MAX_SEARCH = 255

function occurrenceIndex(text: string, needle: string, start: number): number {
  let k = 0
  for (let i = text.indexOf(needle); i !== -1 && i < start; i = text.indexOf(needle, i + 1)) k += 1
  return k
}

function occurrences(text: string, needle: string): number {
  let n = 0
  for (let i = text.indexOf(needle); i !== -1; i = text.indexOf(needle, i + 1)) n += 1
  return n
}

/** Word's Find treats ^ as the start of a special code, so it has to be doubled. */
const escapeFind = (needle: string): string => needle.replace(/\^/g, '^^')

async function paragraphAt(context: Word.RequestContext, index: number, expectedText: string): Promise<Word.Paragraph> {
  const paragraphs = context.document.body.paragraphs
  paragraphs.load('items/text')
  await context.sync()
  const p = paragraphs.items[index]
  if (!p || p.text !== expectedText) throw new StaleDocumentError()
  return p
}

/** The search hit for text[start, end) inside the paragraph, or null when Word's hits do not line up. */
async function rangeFor(context: Word.RequestContext, p: Word.Paragraph, text: string, start: number, end: number): Promise<Word.Range | null> {
  const needle = text.slice(start, end)
  if (!needle || needle.length > MAX_SEARCH) return null
  const hits = p.search(escapeFind(needle), { matchCase: true })
  hits.load('items/text')
  await context.sync()
  if (hits.items.length !== occurrences(text, needle)) return null
  return hits.items[occurrenceIndex(text, needle, start)] ?? null
}

export async function goTo(location: Location): Promise<void> {
  await runWord(async (context) => {
    const body = context.document.body
    switch (location.kind) {
      case 'paragraph': {
        const p = await paragraphAt(context, location.paragraph, location.expectedText)
        const range =
          location.start !== undefined && location.end !== undefined && location.end > location.start
            ? await rangeFor(context, p, location.expectedText, location.start, location.end)
            : null
        if (range) range.select()
        else p.select()
        break
      }
      case 'table': {
        const tables = body.tables
        tables.load('items')
        await context.sync()
        const t = tables.items[location.table]
        if (!t) throw new StaleDocumentError()
        t.select()
        break
      }
      case 'picture': {
        const pictures = body.inlinePictures
        pictures.load('items')
        await context.sync()
        const pic = pictures.items[location.picture]
        if (!pic) throw new StaleDocumentError()
        pic.select()
        break
      }
      case 'control': {
        const controls = body.contentControls
        controls.load('items/tag')
        await context.sync()
        const c = controls.items[location.control]
        if (!c || c.tag !== location.tag) throw new StaleDocumentError()
        c.select()
        break
      }
    }
    await context.sync()
  })
}

type ReplaceFix = Extract<FixAction, { kind: 'replace' }>

export interface Span {
  start: number
  end: number
  replacement: string
}

/**
 * Replaces spans of one paragraph whose current text is `text`. All search hits are found before anything
 * changes, then edits run from the end of the paragraph back, so several spans never collide. Spans whose
 * search hits do not line up with the text are skipped rather than guessed.
 */
export async function replaceSpans(
  context: Word.RequestContext,
  p: Word.Paragraph,
  text: string,
  spans: Span[],
): Promise<{ applied: number; skipped: number }> {
  const usable = spans.filter((s) => s.end > s.start && s.end - s.start <= MAX_SEARCH)
  let skipped = spans.length - usable.length
  let applied = 0
  const needles = [...new Set(usable.map((s) => text.slice(s.start, s.end)))]
  const hits = new Map(needles.map((n) => [n, p.search(escapeFind(n), { matchCase: true })]))
  hits.forEach((h) => h.load('items/text'))
  await context.sync()

  let limit = Number.POSITIVE_INFINITY
  for (const s of [...usable].sort((a, b) => b.start - a.start)) {
    const needle = text.slice(s.start, s.end)
    const found = hits.get(needle)
    const range = found && found.items.length === occurrences(text, needle) ? found.items[occurrenceIndex(text, needle, s.start)] : undefined
    if (!range || s.end > limit) {
      skipped += 1
      continue
    }
    if (s.replacement === '') range.delete()
    else range.insertText(s.replacement, 'Replace')
    limit = s.start
    applied += 1
  }
  await context.sync()
  return { applied, skipped }
}

async function applyReplacements(fixes: ReplaceFix[]): Promise<{ applied: number; skipped: number }> {
  const byParagraph = new Map<number, ReplaceFix[]>()
  for (const f of fixes) byParagraph.set(f.paragraph, [...(byParagraph.get(f.paragraph) ?? []), f])
  const total = { applied: 0, skipped: 0 }
  for (const [index, group] of byParagraph) {
    const first = group[0]
    if (!first) continue
    const text = first.expectedText
    const same = group.filter((f) => f.expectedText === text)
    total.skipped += group.length - same.length
    const r = await runWord(async (context) => replaceSpans(context, await paragraphAt(context, index, text), text, same))
    total.applied += r.applied
    total.skipped += r.skipped
  }
  return total
}

async function ensureStyles(): Promise<void> {
  if ((await hasBwbStyles()) === false) await addBwbStyles()
}

/** Applies several fixes; returns how many worked. Stale paragraphs are skipped, not guessed. */
export async function applyFixes(fixes: FixAction[]): Promise<{ applied: number; skipped: number }> {
  const replacements = fixes.filter((f): f is ReplaceFix => f.kind === 'replace')
  const result = await applyReplacements(replacements)
  // Deleting paragraphs shifts the indexes after them, so work from the end of the document back.
  const others = fixes
    .filter((f) => f.kind !== 'replace')
    .sort((a, b) => ('paragraph' in b ? b.paragraph : -1) - ('paragraph' in a ? a.paragraph : -1))
  for (const f of others) {
    try {
      await applyFix(f)
      result.applied += 1
    } catch (error) {
      if (!(error instanceof StaleDocumentError)) throw error
      result.skipped += 1
    }
  }
  return result
}

export async function applyFix(fix: FixAction): Promise<void> {
  switch (fix.kind) {
    case 'replace': {
      const { applied } = await applyReplacements([fix])
      if (!applied) throw new StaleDocumentError()
      return
    }
    case 'paragraph-style':
      if (!fix.builtIn) await ensureStyles()
      return runWord(async (context) => {
        const p = await paragraphAt(context, fix.paragraph, fix.expectedText)
        if (fix.builtIn) p.styleBuiltIn = fix.builtIn
        else p.style = fix.style
        await context.sync()
      })
    case 'paragraph-font':
      return runWord(async (context) => {
        const p = await paragraphAt(context, fix.paragraph, fix.expectedText)
        p.font.name = fix.font
        await context.sync()
      })
    case 'remove-highlight':
      return runWord(async (context) => {
        const p = await paragraphAt(context, fix.paragraph, fix.expectedText)
        // Office.js removes highlighting when the colour is set to null.
        Object.assign(p.getRange('Whole').font, { highlightColor: null })
        await context.sync()
      })
    case 'delete-paragraph':
      return runWord(async (context) => {
        const p = await paragraphAt(context, fix.paragraph, fix.expectedText)
        p.delete()
        await context.sync()
      })
    case 'table-style':
      await ensureStyles()
      return runWord(async (context) => {
        const tables = context.document.body.tables
        tables.load('items')
        await context.sync()
        const t = tables.items[fix.table]
        if (!t) throw new StaleDocumentError()
        t.style = fix.style
        await context.sync()
      })
    case 'properties':
      return writeDocInfo(fix.values)
    case 'remove-controls':
      return runWord(async (context) => {
        const controls = context.document.body.contentControls.getByTag(fix.tag)
        controls.load('items')
        await context.sync()
        controls.items.forEach((c) => c.delete(false))
        await context.sync()
      })
    case 'remove-sticker':
      await setSticker(null)
      return
    case 'insert-disclaimer':
      await ensureStyles()
      await insertBlocks([{ type: 'disclaimer' }], 'end')
      return
    case 'add-styles':
      await addBwbStyles()
      if (supports('1.5')) await restyleBuiltInStyles()
      return
  }
}
