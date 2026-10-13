// Numbered footnotes on a slide: superscript markers in the text and a footnote block above the source line, one
// note a line as "1 Text". This module plans; footnoteOps.ts does it in PowerPoint. Markers are found by reading
// superscript only on numbers that follow a word or punctuation, so a slide costs one read however much text it has.

import { SOURCE_AREA } from '../../brand/slideGrid.ts'
import { COLOURS } from '../../brand/tokens.ts'
import { splitParagraphs } from '../check/text.ts'
import type { Box, FootnoteMarker, Span } from '../check/types.ts'

/** The tag on the footnote block. */
export const FOOTNOTES_TAG = 'BWB_FOOTNOTES'
export const FOOTNOTE_STYLE = { font: 'Aptos', size: 8, colour: COLOURS.charcoal } as const
/** Points per line of 8 pt text, and the gap kept above the source line. */
const LINE = 10
const GAP = 4
/** About how many characters of 8 pt Aptos fit on a line of the source band. */
const CHARS_PER_LINE = 150

export interface Footnote {
  number: number
  text: string
}

/** The notes in a footnote block. A line that does not start with a number continues the note above it. */
export function parseFootnotes(text: string): Footnote[] {
  const notes: Footnote[] = []
  for (const p of splitParagraphs(text)) {
    const line = p.text.trim()
    if (!line) continue
    const m = /^(\d{1,3})\s+(.*)$/.exec(line) ?? /^(\d{1,3})$/.exec(line)
    const last = notes[notes.length - 1]
    if (m?.[1]) notes.push({ number: Number(m[1]), text: (m[2] ?? '').trim() })
    else if (last) last.text = `${last.text} ${line}`.trim()
  }
  return notes
}

export function formatFootnotes(notes: Footnote[]): string {
  return notes.map((n) => `${n.number} ${n.text}`).join('\n')
}

/** The number for a new footnote: one more than any number in use on the slide. */
export function nextFootnoteNumber(markers: number[], notes: Footnote[]): number {
  return Math.max(0, ...markers, ...notes.map((n) => n.number)) + 1
}

/**
 * Numbers that could be footnote markers: up to three digits straight after a letter, closing bracket, quote or
 * punctuation ("growth3", "2025).4"). Numbers after a space ("USD 63bn", "in 2025") are never markers.
 */
export function markerCandidates(text: string): Span[] {
  const out: Span[] = []
  for (const m of text.matchAll(/(?<=[^\s\d])\d{1,3}(?!\d)/gu)) out.push({ start: m.index, length: m[0].length })
  return out
}

export interface PlacedMarker extends FootnoteMarker {
  path: string[]
  top: number
  left: number
}

/** Shapes whose tops are this close sit on one row and read left to right. */
const ROW = 6

/** Markers in reading order: top to bottom, left to right across shapes, then through each shape's text. */
export function readingOrder(markers: PlacedMarker[]): PlacedMarker[] {
  return [...markers].sort((a, b) => {
    if (a.path.join('/') === b.path.join('/')) return a.start - b.start
    return Math.abs(a.top - b.top) > ROW ? a.top - b.top : a.left - b.left || a.top - b.top
  })
}

export interface RenumberPlan {
  /** Old number to new number, in reading order. */
  mapping: Map<number, number>
  /** Marker text to change, by shape. */
  changes: Array<{ path: string[]; start: number; length: number; text: string }>
  /** The notes renumbered and in order; notes without a marker follow. */
  notes: Footnote[]
}

/** Numbers the markers 1, 2, 3 in reading order (a repeated marker keeps its note), and orders the notes to match. */
export function renumberPlan(markers: PlacedMarker[], notes: Footnote[]): RenumberPlan {
  const mapping = new Map<number, number>()
  const ordered = readingOrder(markers)
  for (const m of ordered) if (!mapping.has(m.number)) mapping.set(m.number, mapping.size + 1)
  const changes = ordered.flatMap((m) => {
    const next = mapping.get(m.number)
    return next === undefined || next === m.number ? [] : [{ path: m.path, start: m.start, length: m.length, text: String(next) }]
  })
  const used = notes.filter((n) => mapping.has(n.number)).map((n) => ({ number: mapping.get(n.number) ?? n.number, text: n.text }))
  used.sort((a, b) => a.number - b.number)
  const unused = notes.filter((n) => !mapping.has(n.number)).map((n, i) => ({ number: mapping.size + i + 1, text: n.text }))
  return { mapping, changes, notes: [...used, ...unused] }
}

/** Notes no marker on the slide refers to. */
export function unusedFootnotes(markers: number[], notes: Footnote[]): Footnote[] {
  const used = new Set(markers)
  return notes.filter((n) => !used.has(n.number))
}

/** Marker numbers with no note. */
export function missingFootnotes(markers: number[], notes: Footnote[]): number[] {
  const have = new Set(notes.map((n) => n.number))
  return [...new Set(markers)].filter((n) => !have.has(n)).sort((a, b) => a - b)
}

/** Where the block goes: across the source band, its bottom just above the source line, growing upwards. */
export function footnoteBox(notes: Footnote[]): Box {
  const lines = Math.max(1, notes.reduce((n, note) => n + Math.max(1, Math.ceil((String(note.number).length + 1 + note.text.length) / CHARS_PER_LINE)), 0))
  const height = lines * LINE + 2
  return { left: SOURCE_AREA.left, top: SOURCE_AREA.top - GAP - height, width: SOURCE_AREA.width, height }
}
