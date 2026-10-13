// Text measurement for laying out chart labels without a browser. Each width is the larger of Aptos and Arial (the
// fallback) for that character, in hundredths of an em, so estimates err on the wide side whichever font draws.

export const FONT_FAMILY = "Aptos, 'Segoe UI', Arial, sans-serif"
export const MIN_FONT_PT = 8
export const LINE_HEIGHT = 1.2
/** The baseline sits this far below the top of a line box, in ems. */
export const BASELINE = 0.86
/** Cap height in ems: a line is centred on a point when its baseline is half of this below it. */
export const CAP_HEIGHT = 0.69
/** Depth of descenders below the baseline, in ems. */
export const DESCENT = 0.22

// Character codes 32 to 126.
const REGULAR = [
  28, 30, 37, 56, 56, 89, 67, 22, 34, 34, 46, 59, 29, 35, 29, 34, 56, 56, 56, 56, 56, 56, 56, 56, 56, 56, 29, 29, 59, 59,
  59, 56, 102, 67, 67, 73, 73, 67, 62, 78, 73, 28, 50, 67, 56, 84, 73, 78, 67, 78, 73, 67, 62, 73, 67, 95, 67, 67, 62, 30,
  34, 30, 54, 56, 56, 56, 57, 53, 57, 56, 31, 56, 56, 24, 24, 50, 27, 86, 56, 56, 57, 57, 34, 50, 33, 56, 50, 73, 50, 50,
  50, 34, 28, 34, 59,
]
const BOLD = [
  28, 34, 48, 56, 56, 89, 73, 24, 34, 34, 46, 59, 30, 35, 30, 39, 56, 56, 56, 56, 56, 56, 56, 56, 56, 56, 34, 34, 59, 59,
  59, 62, 98, 73, 73, 73, 73, 67, 62, 78, 73, 30, 56, 73, 62, 84, 73, 78, 67, 78, 73, 67, 62, 73, 67, 95, 67, 67, 62, 34,
  39, 34, 59, 56, 56, 56, 62, 56, 62, 56, 34, 62, 62, 28, 28, 56, 30, 89, 62, 62, 62, 62, 39, 56, 36, 62, 56, 78, 56, 56,
  50, 39, 28, 39, 59,
]

const SPECIAL: Record<number, number> = {
  0x00a0: 28, 0x2009: 20, 0x2212: 59, 0x2013: 56, 0x2014: 100, 0x2018: 28, 0x2019: 28, 0x201c: 52, 0x201d: 52,
  0x2026: 100, 0x00a3: 56, 0x20ac: 56, 0x00b7: 34, 0x25cf: 75, 0x25cb: 75,
}

function charEm(code: number, bold: boolean): number {
  if (code >= 32 && code < 127) return ((bold ? BOLD : REGULAR)[code - 32] ?? 60) / 100
  const special = SPECIAL[code]
  if (special !== undefined) return special / 100
  if (code >= 0x2e80) return 1
  return bold ? 0.74 : 0.7
}

/** Estimated width of a line of text in points. */
export function textWidth(text: string, size: number, bold = false): number {
  let em = 0
  for (const ch of text) em += charEm(ch.codePointAt(0) ?? 32, bold)
  return em * size
}

export const lineHeight = (size: number): number => size * LINE_HEIGHT

export interface Wrapped {
  lines: string[]
  /** Width of the widest line. */
  width: number
  truncated: boolean
}

/** Shortens text with an ellipsis until it fits; `force` adds the ellipsis even when the text already fits. */
export function ellipsise(text: string, maxWidth: number, size: number, bold = false, force = false): string {
  if (!force && textWidth(text, size, bold) <= maxWidth) return text
  const chars = [...text]
  while (chars.length > 1 && textWidth(`${chars.join('').trimEnd()}\u{2026}`, size, bold) > maxWidth) chars.pop()
  return `${chars.join('').trimEnd()}\u{2026}`
}

interface Token {
  text: string
  space: boolean
}

// Words, with hyphenated words split after the hyphen so "Sub-Saharan" can break as "Sub-" and "Saharan".
function tokens(text: string): Token[] {
  const out: Token[] = []
  for (const word of text.trim().split(/\s+/).filter(Boolean)) {
    word.split(/(?<=[-/])(?=\p{L})/u).forEach((part, i) => out.push({ text: part, space: i === 0 }))
  }
  return out
}

const join = (line: Token[]): string => line.map((t, i) => (i > 0 && t.space ? ` ${t.text}` : t.text)).join('')

/** Breaks text into lines no wider than maxWidth, at most maxLines, ending with an ellipsis if it is cut short. */
export function wrapText(text: string, maxWidth: number, size: number, maxLines = 3, bold = false): Wrapped {
  const lines: Token[][] = []
  let current: Token[] = []
  for (const token of tokens(text)) {
    if (!current.length || textWidth(join([...current, token]), size, bold) <= maxWidth) current.push(token)
    else {
      lines.push(current)
      current = [token]
    }
  }
  if (current.length) lines.push(current)
  let truncated = false
  let out = lines.map(join)
  if (lines.length > maxLines) {
    const rest = join(lines.slice(maxLines - 1).flat())
    out = [...out.slice(0, maxLines - 1), ellipsise(rest, maxWidth, size, bold, true)]
    truncated = true
  }
  const fitted = out.map((line) => {
    if (textWidth(line, size, bold) <= maxWidth) return line
    truncated = true
    return ellipsise(line, maxWidth, size, bold)
  })
  return { lines: fitted, width: Math.max(0, ...fitted.map((l) => textWidth(l, size, bold))), truncated }
}

export interface Fitted {
  size: number
  wrapped: Wrapped[]
  /** The most lines any label takes. */
  lines: number
}

/**
 * Wraps a set of labels at one shared size: the largest from `size` down to the 8 pt minimum at which every label
 * fits its width in `maxLines` lines. Below that, labels are cut short with an ellipsis rather than set smaller.
 */
export function fitLabels(labels: string[], maxWidth: number | number[], size: number, maxLines = 3, bold = false): Fitted {
  const widthOf = (i: number): number => (Array.isArray(maxWidth) ? (maxWidth[i] ?? 0) : maxWidth)
  const wrapAll = (s: number): Wrapped[] => labels.map((label, i) => wrapText(label, widthOf(i), s, maxLines, bold))
  let s = size
  let wrapped = wrapAll(s)
  while (s > MIN_FONT_PT && wrapped.some((w) => w.truncated)) {
    s = Math.max(MIN_FONT_PT, s - 0.5)
    wrapped = wrapAll(s)
  }
  return { size: s, wrapped, lines: Math.max(1, ...wrapped.map((w) => w.lines.length)) }
}
