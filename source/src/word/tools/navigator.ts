// The navigator: the outline of the document (headings, tables and figures, in order), going to an item, promoting
// or demoting the heading at the cursor, and cross-references to a heading, table or figure.

import { LAYOUT_TABLE_STYLE_NAMES } from '../../brand/styles.ts'
import { UserFacingError, runWord } from '../run.ts'
import { captionKind, captionLabel, headingLevel } from './common.ts'
import { unavailable } from './support.ts'

export type OutlineKind = 'heading' | 'table' | 'figure'

export interface OutlineItem {
  /** Unique within one reading of the outline. */
  id: string
  kind: OutlineKind
  /** Heading level, or one deeper than the heading a table or figure sits under. */
  level: number
  label: string
  /** The paragraph to go to and to refer to: the heading, or the caption. -1 for a table with no caption. */
  paragraph: number
  /** That paragraph's text when the outline was read, so going there can check it is still the same. */
  expectedText: string
  /** Position among the document's tables, for tables. */
  table?: number
  /** "Table 3" or "Figure 2", for captions. */
  number?: string
}

export const STALE = 'The document has changed. Refresh the navigator.'

/** Runs of consecutive paragraphs inside tables: one per top-level table, in order. */
function tableRuns(nesting: number[]): number[] {
  const starts: number[] = []
  nesting.forEach((n, i) => {
    if (n > 0 && !((nesting[i - 1] ?? 0) > 0)) starts.push(i)
  })
  return starts
}

/** Reads the outline: headings, data tables (by caption) and figures (by caption), in document order. */
export async function readOutline(): Promise<OutlineItem[]> {
  return runWord(async (context) => {
    const paragraphs = context.document.body.paragraphs
    paragraphs.load('items/text,items/style,items/styleBuiltIn,items/tableNestingLevel')
    const tables = context.document.body.tables
    tables.load('items/style')
    await context.sync()
    const firstCells = tables.items.map((t) => t.getCellOrNullObject(0, 0))
    firstCells.forEach((c) => c.load('value'))
    await context.sync()

    const items = paragraphs.items
    const runs = tableRuns(items.map((p) => p.tableNestingLevel))
    const tableAt = new Map<number, number>()
    if (runs.length === tables.items.length) runs.forEach((start, t) => tableAt.set(start, t))
    const out: OutlineItem[] = []
    const used = new Set<number>()
    let level = 0
    const addTable = (t: number, caption: number): void => {
      const style = tables.items[t]?.style ?? ''
      if (LAYOUT_TABLE_STYLE_NAMES.has(style)) return
      const text = caption >= 0 ? (items[caption]?.text ?? '') : ''
      const first = firstCells[t]
      const cell = first && !first.isNullObject ? first.value.trim() : ''
      if (caption >= 0) used.add(caption)
      const number = caption >= 0 ? captionLabel(text) : null
      out.push({
        id: `table-${t}`, kind: 'table', level: level + 1, table: t, paragraph: caption, expectedText: text,
        label: text.trim() || `Table with no caption${cell ? `: ${cell}` : ''}`, ...(number ? { number } : {}),
      })
    }
    items.forEach((p, i) => {
      const t = tableAt.get(i)
      if (t !== undefined) {
        let end = i
        while ((items[end + 1]?.tableNestingLevel ?? 0) > 0) end += 1
        const after = items[end + 1]
        addTable(t, after && captionKind(after) === 'Table' ? end + 1 : -1)
      }
      if (p.tableNestingLevel > 0 || used.has(i) || !p.text.trim()) return
      const h = headingLevel(p)
      if (h !== null) {
        level = h
        out.push({ id: `heading-${i}`, kind: 'heading', level: h, label: p.text.trim(), paragraph: i, expectedText: p.text })
        return
      }
      const kind = captionKind(p)
      if (!kind) return
      const number = captionLabel(p.text)
      out.push({
        id: `${kind === 'Table' ? 'table' : 'figure'}-p${i}`, kind: kind === 'Table' ? 'table' : 'figure', level: level + 1,
        label: p.text.trim(), paragraph: i, expectedText: p.text, ...(number ? { number } : {}),
      })
    })
    // When the table paragraphs cannot be matched to Word's tables (two tables touching), tables are listed by
    // their captions alone, which the loop above has done.
    return out
  })
}

async function paragraphOf(context: Word.RequestContext, item: OutlineItem): Promise<Word.Paragraph> {
  const paragraphs = context.document.body.paragraphs
  paragraphs.load('items/text')
  await context.sync()
  const p = paragraphs.items[item.paragraph]
  if (!p || p.text !== item.expectedText) throw new UserFacingError(STALE)
  return p
}

/** Moves the cursor to an outline item. */
export async function goToItem(item: OutlineItem): Promise<void> {
  await runWord(async (context) => {
    if (item.kind === 'table' && item.table !== undefined) {
      if (item.paragraph >= 0) await paragraphOf(context, item)
      const tables = context.document.body.tables
      tables.load('items')
      await context.sync()
      const t = tables.items[item.table]
      if (!t) throw new UserFacingError(STALE)
      t.select('Start')
    } else {
      const p = await paragraphOf(context, item)
      p.select('Start')
    }
    await context.sync()
  })
}

const BUILT_IN = ['Heading1', 'Heading2', 'Heading3', 'Heading4'] as const
/** BwB styles headings 1 to 4; its unnumbered headings go to 3. */
const MAX_BUILT_IN = 4
const MAX_UNNUMBERED = 3

/** The heading at the cursor, if the cursor is in one. */
export async function readCursorHeading(): Promise<{ level: number; text: string } | null> {
  return runWord(async (context) => {
    const p = context.document.getSelection().paragraphs.getFirst()
    p.load('text,style,styleBuiltIn')
    await context.sync()
    const level = headingLevel(p)
    return level === null ? null : { level, text: p.text.trim() }
  })
}

/** Moves the headings in the selection one level up (-1) or down (+1). */
export async function shiftHeadingLevel(delta: 1 | -1): Promise<string> {
  return runWord(async (context) => {
    const paragraphs = context.document.getSelection().paragraphs
    paragraphs.load('items/text,items/style,items/styleBuiltIn')
    await context.sync()
    const headings = paragraphs.items.map((p) => ({ p, level: headingLevel(p), unnumbered: /Unnumbered$/.test(p.style) }))
      .filter((h): h is { p: Word.Paragraph; level: number; unnumbered: boolean } => h.level !== null)
    if (!headings.length) throw new UserFacingError('Put the cursor in a heading first.')
    const moves = headings.filter((h) => {
      const next = h.level + delta
      return next >= 1 && next <= (h.unnumbered ? MAX_UNNUMBERED : MAX_BUILT_IN)
    })
    if (!moves.length) {
      throw new UserFacingError(delta < 0 ? 'Heading 1 is the top level already.' : `BwB headings go down to Heading ${headings.every((h) => h.unnumbered) ? MAX_UNNUMBERED : MAX_BUILT_IN}.`)
    }
    for (const h of moves) {
      const next = h.level + delta
      if (h.unnumbered) h.p.style = `BwB Heading ${next} Unnumbered`
      else h.p.styleBuiltIn = BUILT_IN[next - 1] ?? 'Heading4'
    }
    await context.sync()
    const [only] = moves
    if (moves.length === 1 && only) return `Heading ${only.level} is now Heading ${only.level + delta}.`
    return `${moves.length} headings ${delta < 0 ? 'promoted' : 'demoted'} one level.`
  })
}

/** A bookmark name Word accepts: a letter first, then letters, digits and underscores, at most 40 characters. */
export function bookmarkBase(label: string): string {
  const words = label.replace(/[^A-Za-z0-9]+/g, ' ').trim().split(' ').filter(Boolean)
  return `BwB_${words.join('_')}`.slice(0, 36).replace(/_+$/, '')
}

/**
 * Inserts a cross-reference (a REF field with a link) at the cursor to a heading, table or figure, bookmarking the
 * target first if it has no bookmark. A heading shows its text, or its number with `show: 'number'`; a caption
 * shows its label, such as Table 3. Word on the web cannot insert fields, so this needs Word for Windows or Mac.
 */
export async function insertCrossReference(item: OutlineItem, show: 'text' | 'number' = 'text'): Promise<string> {
  const reason = unavailable('crossReference')
  if (reason) throw new UserFacingError(reason)
  if (item.paragraph < 0) throw new UserFacingError('This table has no caption to refer to. Add a caption below it first.')
  return runWord(async (context) => {
    const p = await paragraphOf(context, item)
    const hits = item.kind !== 'heading' && item.number ? p.search(item.number, { matchCase: true }) : null
    hits?.load('items/text')
    const content = p.getRange('Content')
    content.load('text')
    await context.sync()
    const target = hits?.items[0] ?? content
    target.load('text')
    const existing = target.getBookmarks(true, false)
    await context.sync()
    const ranges = existing.value.map((name) => ({ name, range: context.document.getBookmarkRangeOrNullObject(name) }))
    ranges.forEach((r) => r.range.load('text'))
    const base = bookmarkBase(item.number ?? item.label)
    const candidates = [base, ...Array.from({ length: 8 }, (_, i) => `${base}_${i + 2}`)]
    const taken = candidates.map((name) => ({ name, range: context.document.getBookmarkRangeOrNullObject(name) }))
    taken.forEach((t) => t.range.load('isNullObject'))
    await context.sync()
    let name = ranges.find((r) => !r.range.isNullObject && r.range.text.trim() === target.text.trim())?.name
    if (!name) {
      name = taken.find((t) => t.range.isNullObject)?.name
      if (!name) throw new UserFacingError('Could not name a bookmark for this target. Insert the cross-reference from Word: References, Cross-reference.')
      target.insertBookmark(name)
    }
    const field = context.document.getSelection().insertField('After', 'Ref', `${name}${show === 'number' && item.kind === 'heading' ? ' \\r' : ''} \\h`)
    field.updateResult()
    await context.sync()
    const what = item.kind === 'heading' ? `\u201c${item.label}\u201d` : (item.number ?? `\u201c${item.label}\u201d`)
    return `Cross-reference to ${what} inserted. It updates with the document\u2019s fields (F9).`
  })
}
