// The BwB template's styles (BwB Document Format skill v1.1.0, references/style-map.md).
// `name` is the style name Word shows and Office.js uses; built-in styles also carry their
// locale-independent `builtIn` identifier so they work in any Word language.

export type BuiltIn = 'Normal' | 'Heading1' | 'Heading2' | 'Heading3' | 'Heading4' | 'Title' | 'Subtitle' | 'Caption'

export interface ParagraphStyle {
  id: string
  name: string
  builtIn?: BuiltIn
  label: string
  hint: string
  preview: { font: 'Candara' | 'Aptos'; sizePt: number; colour: string; bold?: boolean; italic?: boolean }
}

const NAVY = '#022338'
const CHARCOAL = '#4A4A4C'
const BLACK = '#000000'

export const HEADING_STYLES: ParagraphStyle[] = [
  { id: 'Heading1', name: 'heading 1', builtIn: 'Heading1', label: 'Heading 1', hint: 'Numbered 1. Candara 18 pt', preview: { font: 'Candara', sizePt: 18, colour: NAVY, bold: true } },
  { id: 'Heading2', name: 'heading 2', builtIn: 'Heading2', label: 'Heading 2', hint: 'Numbered 1.1 Candara 14 pt', preview: { font: 'Candara', sizePt: 14, colour: NAVY, bold: true } },
  { id: 'Heading3', name: 'heading 3', builtIn: 'Heading3', label: 'Heading 3', hint: 'Numbered 1.1.1 Candara 12 pt', preview: { font: 'Candara', sizePt: 12, colour: NAVY, bold: true } },
  { id: 'Heading4', name: 'heading 4', builtIn: 'Heading4', label: 'Heading 4', hint: 'Not numbered. Candara 11 pt bold italic', preview: { font: 'Candara', sizePt: 11, colour: NAVY, bold: true, italic: true } },
]

export const UNNUMBERED_HEADING_STYLES: ParagraphStyle[] = [
  { id: 'BwBHeading1Unnumbered', name: 'BwB Heading 1 Unnumbered', label: 'Heading 1, not numbered', hint: 'Executive summary, glossary, references, annexes', preview: { font: 'Candara', sizePt: 18, colour: NAVY, bold: true } },
  { id: 'BwBHeading2Unnumbered', name: 'BwB Heading 2 Unnumbered', label: 'Heading 2, not numbered', hint: 'Sub-sections of front and back matter', preview: { font: 'Candara', sizePt: 14, colour: NAVY, bold: true } },
  { id: 'BwBHeading3Unnumbered', name: 'BwB Heading 3 Unnumbered', label: 'Heading 3, not numbered', hint: 'Sub-sections of front and back matter', preview: { font: 'Candara', sizePt: 12, colour: NAVY, bold: true } },
]

export const TEXT_STYLES: ParagraphStyle[] = [
  { id: 'Normal', name: 'Normal', builtIn: 'Normal', label: 'Body text', hint: 'Aptos 10 pt, justified', preview: { font: 'Aptos', sizePt: 10, colour: BLACK } },
  { id: 'BwBLead', name: 'BwB Lead', label: 'Lead paragraph', hint: 'An introductory sentence under a heading. Use sparingly', preview: { font: 'Aptos', sizePt: 12, colour: CHARCOAL, italic: true } },
  { id: 'BwBAccent', name: 'BwB Accent Paragraph', label: 'Accent paragraph', hint: 'One highlighted sentence with a green bar. At most twice a page', preview: { font: 'Aptos', sizePt: 11, colour: NAVY, bold: true } },
  { id: 'BwBSource', name: 'BwB Source', label: 'Source line', hint: 'A free-standing source or note. Aptos 8 pt italic', preview: { font: 'Aptos', sizePt: 8, colour: CHARCOAL, italic: true } },
  { id: 'BwBPullQuote', name: 'BwB Pull Quote', label: 'Pull quote', hint: 'One short quotation that carries a point', preview: { font: 'Candara', sizePt: 16, colour: NAVY, italic: true } },
]

export interface ListStyle {
  id: string
  name: string
  label: string
  hint: string
  marker: string
  markerColour: string
  onNavy?: boolean
  numbered?: boolean
}

export const LIST_STYLES: ListStyle[] = [
  { id: 'BwBBulletGreenSquare', name: 'BwB Bullet Green Square', label: 'Green square', hint: 'Body lists', marker: '■', markerColour: '#328840' },
  { id: 'BwBBulletNavyCircle', name: 'BwB Bullet Navy Circle', label: 'Navy circle', hint: 'Secondary lists', marker: '●', markerColour: NAVY },
  { id: 'BwBBulletGreenCircle', name: 'BwB Bullet Green Circle', label: 'Green circle', hint: 'Inside callouts and tables', marker: '●', markerColour: '#328840' },
  { id: 'BwBBulletCharcoalDash', name: 'BwB Bullet Charcoal Dash', label: 'Charcoal dash', hint: 'Assumptions and minor items', marker: '–', markerColour: CHARCOAL },
  { id: 'BwBBulletArrow', name: 'BwB Bullet Arrow', label: 'Green arrow', hint: 'Actions and next steps', marker: '›', markerColour: '#328840' },
  { id: 'BwBNumberedNavy', name: 'BwB Numbered Navy', label: 'Numbered', hint: 'Sequence or ranking', marker: '1.', markerColour: NAVY, numbered: true },
  { id: 'BwBBulletLimeNavy', name: 'BwB Bullet Lime on Navy', label: 'Lime on navy', hint: 'Only inside navy panels', marker: '■', markerColour: '#BED62F', onNavy: true },
]

export const CAPTION_STYLES = {
  table: { id: 'BwBTableCaption', name: 'BwB Table Caption' },
  figure: { id: 'BwBImageCaption', name: 'BwB Image Caption' },
  label: { id: 'BwBCaptionLabel', name: 'BwB Caption Label' },
} as const

export const TABLE_STYLES = {
  standard: { id: 'BwBStandard', name: 'BwB Standard', label: 'Standard data table' },
  comparison: { id: 'BwBComparison', name: 'BwB Comparison', label: 'Comparison (first column is a term)' },
  status: { id: 'BwBCountryStatus', name: 'BwB Country Status', label: 'Status table' },
} as const

/** Data table styles the checker accepts; layout tables (panels, callouts) use the others. */
export const DATA_TABLE_STYLE_NAMES = new Set(['BwB Standard', 'BwB Comparison', 'BwB Country Status'])
export const LAYOUT_TABLE_STYLE_NAMES = new Set([
  'BwB Snapshot', 'BwB Highlight', 'BwB Highlight Charcoal', 'BwB Panel (navy)', 'BwB Plain (layout)',
])

/** Custom styles whose presence shows a document is built on the BwB template. */
export const TEMPLATE_MARKER_STYLES = ['BwB Lead', 'BwB Bullet Green Square', 'BwB Table Caption', 'BwB Callout Title', 'BwB Standard']

export function isHeadingStyleName(name: string): number | null {
  const m = /^(?:heading ([1-6])|BwB Heading ([1-3]) Unnumbered)$/i.exec(name.trim())
  if (!m) return null
  return Number(m[1] ?? m[2])
}

export function isCaptionStyleName(name: string): 'table' | 'figure' | null {
  if (name === CAPTION_STYLES.table.name) return 'table'
  if (name === CAPTION_STYLES.figure.name) return 'figure'
  if (name.toLowerCase() === 'caption') return 'figure'
  return null
}
