// Text helpers for slides. PowerPoint separates paragraphs with \r (some builds \n) and lines with a vertical tab,
// and its TextRange has no runs, so the reader splits the text into paragraphs and words and reads each piece.

import type { Span } from './types.ts'

const BREAK = /\r\n|[\r\n\v\f\u2028\u2029]/g

/** Every paragraph or line of `text`, with its offsets. Empty ones are kept, so callers can count lines. */
export function splitParagraphs(text: string): Array<Span & { text: string }> {
  const out: Array<Span & { text: string }> = []
  let start = 0
  for (const m of text.matchAll(BREAK)) {
    out.push({ start, length: m.index - start, text: text.slice(start, m.index) })
    start = m.index + m[0].length
  }
  out.push({ start, length: text.length - start, text: text.slice(start) })
  return out
}

/** The runs of non-space characters in text[start, start + length), with offsets into `text`. */
export function splitWords(text: string, start = 0, length = text.length - start): Span[] {
  const out: Span[] = []
  const slice = text.slice(start, start + length)
  for (const m of slice.matchAll(/\S+/g)) out.push({ start: start + m.index, length: m[0].length })
  return out
}

/** Text on one line, as a title reads in a list: breaks become spaces, runs of spaces collapse. */
export function oneLine(text: string): string {
  return text.replace(BREAK, ' ').replace(/[\s\u00A0]+/g, ' ').trim()
}

export function wordCount(text: string): number {
  const words = text.trim().match(/\S+/g)
  return words ? words.length : 0
}

/** Upper-case hex without '#' for '#rrggbb' or 'rrggbb'; null for anything else (mixed, theme names, empty). */
export function hexColour(value: string | null | undefined): string | null {
  if (!value) return null
  const m = /^#?([0-9a-f]{6})$/i.exec(value.trim())
  return m?.[1] ? m[1].toUpperCase() : null
}

/** A point size for messages: 9, 9.5, 10.25. */
export function points(size: number): string {
  return `${Number.isInteger(size) ? size : Math.round(size * 100) / 100} pt`
}

/** A section name short enough for a tracker: cut at a word boundary, with an ellipsis. */
export function shorten(text: string, max: number): string {
  const line = oneLine(text)
  if (line.length <= max) return line
  const cut = line.slice(0, max - 1)
  const space = cut.lastIndexOf(' ')
  return `${(space > max / 2 ? cut.slice(0, space) : cut).replace(/[\s,;:.]+$/, '')}…`
}

/** Merges touching or overlapping spans and sorts them, so a fix sets each stretch once. */
export function mergeSpans(spans: Span[]): Span[] {
  const sorted = [...spans].filter((s) => s.length > 0).sort((a, b) => a.start - b.start)
  const out: Span[] = []
  for (const s of sorted) {
    const last = out[out.length - 1]
    if (last && s.start <= last.start + last.length) {
      last.length = Math.max(last.length, s.start + s.length - last.start)
    } else {
      out.push({ ...s })
    }
  }
  return out
}
