// SVG markup helpers. Every chart is drawn in points: the viewBox is the chart's size in points, so a font size of 8
// is 8 pt once the chart is inserted at that size. Attributes rather than CSS, because PowerPoint's SVG import and
// "Convert to Shape" read presentation attributes most reliably.

import { BASELINE, CAP_HEIGHT, FONT_FAMILY, LINE_HEIGHT } from './measure.ts'

// Characters XML 1.0 does not allow, even escaped: C0 controls other than tab, line feed and carriage return.
const INVALID_XML = /[\u{0}-\u{8}\u{B}\u{C}\u{E}-\u{1F}\u{FFFE}\u{FFFF}]/gu

/** Escapes text for use in element content and attribute values. */
export function esc(text: string): string {
  return text
    .replace(INVALID_XML, '')
    .replace(/&/g, '&amp;')
    .replace(/</g, '&lt;')
    .replace(/>/g, '&gt;')
    .replace(/"/g, '&quot;')
    .replace(/'/g, '&apos;')
}

/** A coordinate rounded to two decimals, never "-0". */
export function num(value: number): string {
  const rounded = Math.round(value * 100) / 100
  return Object.is(rounded, -0) || !Number.isFinite(rounded) ? '0' : String(rounded)
}

export type Anchor = 'start' | 'middle' | 'end'

export interface TextStyle {
  size: number
  fill: string
  anchor?: Anchor
  bold?: boolean
  italic?: boolean
}

/** One line of text with its baseline at y. */
export function text(x: number, y: number, content: string, style: TextStyle): string {
  const attrs = [
    `x="${num(x)}"`,
    `y="${num(y)}"`,
    `font-family="${esc(FONT_FAMILY)}"`,
    `font-size="${num(style.size)}"`,
    `fill="#${style.fill}"`,
  ]
  if (style.anchor && style.anchor !== 'start') attrs.push(`text-anchor="${style.anchor}"`)
  if (style.bold) attrs.push('font-weight="bold"')
  if (style.italic) attrs.push('font-style="italic"')
  return `<text ${attrs.join(' ')}>${esc(content)}</text>`
}

/** Baseline that centres a line of text on y. */
export const centredBaseline = (y: number, size: number): number => y + (CAP_HEIGHT * size) / 2

/** Lines of text whose block starts at top; one text element per line so each converts to its own shape. */
export function textLines(x: number, top: number, lines: string[], style: TextStyle): string {
  return lines.map((line, i) => text(x, top + (BASELINE + i * LINE_HEIGHT) * style.size, line, style)).join('')
}

/** Lines of text centred vertically on y. */
export function textLinesCentred(x: number, y: number, lines: string[], style: TextStyle): string {
  const blockHeight = (lines.length - 1) * LINE_HEIGHT * style.size
  return lines
    .map((line, i) => text(x, centredBaseline(y - blockHeight / 2 + i * LINE_HEIGHT * style.size, style.size), line, style))
    .join('')
}

export function rect(x: number, y: number, w: number, h: number, fill: string): string {
  return `<rect x="${num(x)}" y="${num(y)}" width="${num(Math.max(0, w))}" height="${num(Math.max(0, h))}" fill="#${fill}"/>`
}

export function line(x1: number, y1: number, x2: number, y2: number, stroke: string, width: number, dash?: string): string {
  const dashed = dash ? ` stroke-dasharray="${dash}"` : ''
  return `<line x1="${num(x1)}" y1="${num(y1)}" x2="${num(x2)}" y2="${num(y2)}" stroke="#${stroke}" stroke-width="${num(width)}"${dashed}/>`
}

export function path(d: string, fill: string, stroke?: { colour: string; width: number }): string {
  const outline = stroke ? ` stroke="#${stroke.colour}" stroke-width="${num(stroke.width)}" stroke-linejoin="round"` : ''
  return `<path d="${d}" fill="${fill === 'none' ? 'none' : `#${fill}`}"${outline}/>`
}

export function circle(cx: number, cy: number, r: number, fill: string, stroke?: { colour: string; width: number }): string {
  const outline = stroke ? ` stroke="#${stroke.colour}" stroke-width="${num(stroke.width)}"` : ''
  return `<circle cx="${num(cx)}" cy="${num(cy)}" r="${num(r)}" fill="${fill === 'none' ? 'none' : `#${fill}`}"${outline}/>`
}

/** A diamond (a milestone or a legend mark) centred on cx, cy. */
export function diamond(cx: number, cy: number, half: number, fill: string): string {
  const d = `M${num(cx)} ${num(cy - half)}L${num(cx + half)} ${num(cy)}L${num(cx)} ${num(cy + half)}L${num(cx - half)} ${num(cy)}Z`
  return path(d, fill)
}

/** Groups markup under a label, so a chart converted to shapes in PowerPoint keeps named groups. */
export const group = (id: string, content: string): string => (content ? `<g id="${esc(id)}">${content}</g>` : '')

/**
 * The finished document. The viewBox is in points; width and height are the same size in CSS pixels (four thirds of
 * a point), so anything that sizes the picture from the file gets its true size.
 */
export function svgDocument(widthPt: number, heightPt: number, body: string, label: string): string {
  const px = (pt: number): string => num((pt * 4) / 3)
  return (
    `<svg xmlns="http://www.w3.org/2000/svg" width="${px(widthPt)}" height="${px(heightPt)}" ` +
    `viewBox="0 0 ${num(widthPt)} ${num(heightPt)}" role="img" aria-label="${esc(label)}">${body}</svg>`
  )
}
