// Shared helpers for the Word power tools: units, counts in messages, captions, headings and mark-up escaping.

import { CAPTION_STYLES, isHeadingStyleName } from '../../brand/styles.ts'

export const CM_TO_PT = 72 / 2.54
/** The BwB A4 text width: 21 cm less two 2.2 cm margins. */
export const TEXT_WIDTH_CM = 16.6

export const cmToPt = (cm: number): number => cm * CM_TO_PT
export const ptToCm = (pt: number): number => pt / CM_TO_PT

/** 16.6 for 16.6, 4 for 4.0. */
export const formatCm = (cm: number): string => String(Math.round(cm * 10) / 10)

/** "1 picture", "3 pictures". */
export function count(n: number, singular: string, plural = `${singular}s`): string {
  return `${n} ${n === 1 ? singular : plural}`
}

/** "A", "A and B", "A, B and C". */
export function listOf(items: string[]): string {
  if (items.length < 2) return items.join('')
  return `${items.slice(0, -1).join(', ')} and ${items[items.length - 1] ?? ''}`
}

/** No text a reader sees: a picture's line has none, nor do Word's control characters. */
export const isBlank = (text: string): boolean => /^[\s\u0000-\u001f\u00a0\u200b]*$/.test(text)

export interface ParagraphLike {
  text: string
  style: string
  styleBuiltIn?: string
}

// A label then a full stop, colon or dash ("Figure 3." or "Table 2:"), so "Figure 3 shows ..." in the text is not one.
const FIGURE_LABEL = /^\s*(?:Figure|Chart|Graph|Map|Exhibit|Image|Photo)\s+\d+(?:[.-]\d+)*\s*(?:[.:\u2013-]|$)/i
const TABLE_LABEL = /^\s*Table\s+\d+(?:[.-]\d+)*\s*(?:[.:\u2013-]|$)/i
const LABEL = /^\s*((?:Figure|Chart|Graph|Map|Exhibit|Image|Photo|Table)\s+\d+(?:[.-]\d+)*)/i

/** Whether a paragraph is a table or figure caption, by its BwB style or its "Figure 3" label. */
export function captionKind(p: ParagraphLike): 'Table' | 'Figure' | null {
  if (p.style === CAPTION_STYLES.table.name) return 'Table'
  if (p.style === CAPTION_STYLES.figure.name) return 'Figure'
  if (TABLE_LABEL.test(p.text)) return 'Table'
  if (FIGURE_LABEL.test(p.text)) return 'Figure'
  if (p.styleBuiltIn === 'Caption' || p.style.toLowerCase() === 'caption') return 'Figure'
  return null
}

/** "Table 3" from "Table 3. Green bond issuance", or null. */
export const captionLabel = (text: string): string | null => LABEL.exec(text)?.[1]?.replace(/\s+/, ' ') ?? null

/** The title of a caption: no "Figure 3." label, no "Source: ..." and nothing when it is still a placeholder. */
export function captionTitle(text: string): string {
  const title = text
    .replace(/^\s*(?:Figure|Chart|Graph|Map|Exhibit|Image|Photo|Table)\s+\d+(?:[.-]\d+)*\s*[.:\u2013-]?\s*/i, '')
    .replace(/\s*\bSources?:[\s\S]*$/i, '')
    .trim()
  return /^\[[^\]]*\]\.?$/.test(title) ? '' : title
}

/** Heading level 1 to 9 from a built-in heading style or a BwB unnumbered heading style, or null. */
export function headingLevel(p: { style: string; styleBuiltIn: string }): number | null {
  const builtIn = /^Heading([1-9])$/.exec(p.styleBuiltIn)
  if (builtIn) return Number(builtIn[1])
  return isHeadingStyleName(p.style)
}

/** Escapes the builder's inline mark-up (* [ ] ^ \) so text from the document is inserted as it reads. */
export const escapeMarkup = (text: string): string => text.replace(/[\\*[\]^]/g, (c) => `\\${c}`)
