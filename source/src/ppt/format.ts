// BwB formatting for the selected shapes: fills, text colours, outlines, fonts, text frames, and the house-style
// wording clean-up. Groups are opened up into their members (PowerPointApi 1.8).

import { COLOURS, FONTS } from '../brand/tokens.ts'
import { cmToPt } from '../geometry/index.ts'
import { quoteIssues, scan } from '../rules/writing.ts'
import { UserFacingError } from '../word/run.ts'
import { requirePpt, runPpt, supportsPpt } from './run.ts'
import { RULE } from './specs.ts'

/** Shape types that hold text. */
const TEXT_TYPES = new Set(['GeometricShape', 'TextBox', 'Placeholder', 'Callout'])
const TITLES = new Set(['Title', 'CenterTitle', 'VerticalTitle'])

const plural = (n: number, word: string): string => `${n} ${word}${n === 1 ? '' : 's'}`

interface Leaf {
  shape: PowerPoint.Shape
  name: string
  type: string
}

/** The selected shapes, with groups (and groups in groups) replaced by their members when PowerPoint allows. */
async function selectedLeaves(context: PowerPoint.RequestContext): Promise<Leaf[]> {
  const top = context.presentation.getSelectedShapes()
  top.load('items/id,items/name,items/type')
  await context.sync()
  const openGroups = supportsPpt('1.8')
  const out: Leaf[] = []
  let level: PowerPoint.Shape[] = top.items
  for (let depth = 0; level.length && depth < 8; depth += 1) {
    const groups = openGroups ? level.filter((s) => s.type === 'Group') : []
    out.push(...level.filter((s) => !groups.includes(s)).map((shape) => ({ shape, name: shape.name, type: shape.type })))
    if (!groups.length) break
    const members = groups.map((g) => {
      const shapes = g.group.shapes
      shapes.load('items/id,items/name,items/type')
      return shapes
    })
    await context.sync()
    level = members.flatMap((m) => m.items)
  }
  return out
}

async function onSelection<T>(fn: (context: PowerPoint.RequestContext, leaves: Leaf[]) => Promise<T>): Promise<T> {
  requirePpt('1.5', 'Formatting shapes')
  return runPpt(async (context) => {
    const leaves = await selectedLeaves(context)
    if (!leaves.length) throw new UserFacingError('Select the shapes to format first.')
    return fn(context, leaves)
  })
}

/** Runs `apply` on every selected shape that holds text; `done` names what happened, as in "Autofit off for". */
async function onText(apply: (frame: PowerPoint.TextFrame) => void, done: string): Promise<string> {
  return onSelection(async (context, leaves) => {
    const texts = leaves.filter((l) => TEXT_TYPES.has(l.type))
    if (!texts.length) throw new UserFacingError('Select shapes or text boxes first. Pictures, lines and tables have no text frame.')
    texts.forEach((l) => apply(l.shape.textFrame))
    await context.sync()
    return `${done} ${plural(texts.length, 'shape')}.`
  })
}

export const FILLS = {
  navy: { fill: COLOURS.navy, text: COLOURS.white, label: 'Navy fill, white text,' },
  green: { fill: COLOURS.green, text: COLOURS.white, label: 'Green fill, white text,' },
  offWhite: { fill: COLOURS.offWhite, text: COLOURS.navy, label: 'Off-white fill, navy text,' },
  none: { fill: null, text: null, label: 'No fill' },
} as const
export type FillChoice = keyof typeof FILLS

/** A BwB fill, with the text colour that goes with it. */
export async function applyFill(choice: FillChoice): Promise<string> {
  const { fill, text, label } = FILLS[choice]
  return onSelection(async (context, leaves) => {
    const targets = leaves.filter((l) => TEXT_TYPES.has(l.type))
    if (!targets.length) throw new UserFacingError('Select shapes or text boxes to fill. Pictures, lines and tables keep their own.')
    for (const l of targets) {
      if (fill) l.shape.fill.setSolidColor(`#${fill}`)
      else l.shape.fill.clear()
      if (text) l.shape.textFrame.textRange.font.color = `#${text}`
    }
    await context.sync()
    return `${label} on ${plural(targets.length, 'shape')}.`
  })
}

export const TEXT_COLOURS = { navy: COLOURS.navy, charcoal: COLOURS.charcoal, white: COLOURS.white } as const
export type TextColour = keyof typeof TEXT_COLOURS

export async function applyTextColour(colour: TextColour): Promise<string> {
  const note = colour === 'white' ? ' Use white on navy or green only.' : ''
  const done = await onText((frame) => {
    frame.textRange.font.color = `#${TEXT_COLOURS[colour]}`
  }, `${colour.charAt(0).toUpperCase()}${colour.slice(1)} text in`)
  return `${done}${note}`
}

/** A light grey 0.75 pt outline, or none. Lines keep theirs. */
export async function applyOutline(choice: 'grey' | 'none'): Promise<string> {
  return onSelection(async (context, leaves) => {
    const targets = leaves.filter((l) => TEXT_TYPES.has(l.type) || l.type === 'Image')
    if (!targets.length) throw new UserFacingError('Select shapes, text boxes or pictures first.')
    for (const l of targets) {
      const line = l.shape.lineFormat
      if (choice === 'none') {
        line.visible = false
        continue
      }
      line.visible = true
      line.color = `#${RULE.colour}`
      line.weight = RULE.weight
      line.dashStyle = 'Solid'
    }
    await context.sync()
    return `${choice === 'grey' ? 'Light grey outline on' : 'No outline on'} ${plural(targets.length, 'shape')}.`
  })
}

/** Which of these text shapes are title placeholders: by placeholder type (1.8), or by name in older PowerPoint. */
async function titleFlags(context: PowerPoint.RequestContext, texts: readonly Leaf[]): Promise<boolean[]> {
  const typed = supportsPpt('1.8')
  const placeholders = texts.map((l) => {
    if (l.type !== 'Placeholder' || !typed) return null
    const format = l.shape.placeholderFormat
    format.load('type')
    return format
  })
  if (placeholders.some(Boolean)) await context.sync()
  return texts.map((l, i) => l.type === 'Placeholder' && (typed ? TITLES.has(placeholders[i]?.type ?? '') : /^(center )?title\b/i.test(l.name)))
}

/** Candara for title placeholders, Aptos for everything else; with PowerPointApi 1.9, table cells too. */
export async function applyBwbFonts(): Promise<string> {
  return onSelection(async (context, leaves) => {
    const texts = leaves.filter((l) => TEXT_TYPES.has(l.type))
    const tables = supportsPpt('1.9') ? leaves.filter((l) => l.type === 'Table') : []
    if (!texts.length && !tables.length) throw new UserFacingError('Select shapes or text boxes first.')
    const titles = await titleFlags(context, texts)
    texts.forEach((l, i) => (l.shape.textFrame.textRange.font.name = titles[i] ? FONTS.heading : FONTS.body))
    const loaded = tables.map((l) => {
      const table = l.shape.getTable()
      table.load('rowCount,columnCount')
      return table
    })
    await context.sync()
    for (const table of loaded) {
      for (let r = 0; r < table.rowCount; r += 1) {
        for (let c = 0; c < table.columnCount; c += 1) table.getCellOrNullObject(r, c).font.name = FONTS.body
      }
    }
    if (loaded.length) await context.sync()
    const n = titles.filter(Boolean).length
    const parts = [
      n ? `Candara on ${plural(n, 'title')}` : '',
      texts.length - n ? `Aptos on ${plural(texts.length - n, 'shape')}` : '',
      loaded.length ? `Aptos in ${plural(loaded.length, 'table')}` : '',
    ]
    return `${parts.filter(Boolean).join(', ')}.`
  })
}

/** BwB standard text margins: 0.2 cm on every side. */
export const STANDARD_MARGIN = cmToPt(0.2)

export async function setTextMargins(kind: 'zero' | 'standard'): Promise<string> {
  const m = kind === 'zero' ? 0 : STANDARD_MARGIN
  return onText((frame) => {
    frame.leftMargin = m
    frame.rightMargin = m
    frame.topMargin = m
    frame.bottomMargin = m
  }, kind === 'zero' ? 'No text margins in' : 'Standard 0.2 cm text margins in')
}

export const setAutofitOff = (): Promise<string> => onText((frame) => (frame.autoSizeSetting = 'AutoSizeNone'), 'Autofit off for')

export const setWordWrap = (on: boolean): Promise<string> => onText((frame) => (frame.wordWrap = on), on ? 'Text wraps in' : 'Text runs on one line in')

export type VerticalAnchor = 'Top' | 'Middle' | 'Bottom'

export const setVerticalAnchor = (anchor: VerticalAnchor): Promise<string> =>
  onText((frame) => (frame.verticalAlignment = anchor), `Text anchored at the ${anchor.toLowerCase()} of`)

export interface Fix {
  start: number
  end: number
  replacement: string
}

/** The safe house-style fixes in a text, found paragraph by paragraph, as spans of the whole text. */
export function wordingFixes(text: string, title = false): Fix[] {
  const parts = text.split(/(\r\n|[\r\n\v\u2028\u2029])/)
  const out: Fix[] = []
  let at = 0
  parts.forEach((part, i) => {
    if (i % 2 === 0 && part.trim()) {
      for (const issue of [...scan(part, { title }), ...quoteIssues(part)]) {
        if (issue.fixable && issue.replacement !== undefined) out.push({ start: at + issue.start, end: at + issue.end, replacement: issue.replacement })
      }
    }
    at += part.length
  })
  return out
}

/**
 * Fixes the house style that has one right answer in the selected shapes' text: UK spelling, dashes, dates,
 * currency codes, per cent, the company name and quotation marks. Each fix replaces only its own characters, from
 * the end of the text back, so the formatting around it stays.
 */
export async function cleanSelectedText(): Promise<string> {
  return onSelection(async (context, leaves) => {
    const texts = leaves.filter((l) => TEXT_TYPES.has(l.type))
    if (!texts.length) throw new UserFacingError('Select shapes or text boxes with text first.')
    const ranges = texts.map((l) => {
      const range = l.shape.textFrame.textRange
      range.load('text')
      return range
    })
    const titles = await titleFlags(context, texts)
    await context.sync()
    let applied = 0
    let skipped = 0
    ranges.forEach((range, i) => {
      let limit = Number.POSITIVE_INFINITY
      for (const fix of wordingFixes(range.text, titles[i]).sort((a, b) => b.start - a.start || b.end - a.end)) {
        if (fix.end > limit || fix.end <= fix.start) {
          skipped += 1
          continue
        }
        range.getSubstring(fix.start, fix.end - fix.start).text = fix.replacement
        limit = fix.start
        applied += 1
      }
    })
    if (applied) await context.sync()
    if (!applied && !skipped) return 'Nothing to change in the selection.'
    const changed = `${applied} ${applied === 1 ? 'change' : 'changes'} made`
    return skipped ? `${changed}; ${skipped} left for you to check.` : `${changed}.`
  })
}
