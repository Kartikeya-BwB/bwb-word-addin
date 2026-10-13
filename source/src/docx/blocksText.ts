/** Text blocks: paragraphs, headings, lists, captions, pull quotes, figures and the guidance prompt (builder.Doc). */
import type { Doc } from './builder.ts'
import { inlinePicture } from './drawing.ts'
import { fieldsAsText } from './inline.ts'
import { stripIconTokens } from './iconPolicy.ts'
import type { FigureBlock, ListItem } from './types.ts'
import {
  cpHead,
  EMU_PER_TWIP,
  field,
  para,
  pyRstrip,
  pyStrip,
  pyStripChars,
  ppr,
  run,
  spacer,
  TEXT_W,
  type ParaProps,
} from './xml.ts'

// numbering ids (contract with the template: references/style-map.md)
export const NUM_GREEN_SQUARE = 2
export const NUM_NAVY_CIRCLE = 3
export const NUM_GREEN_CIRCLE = 4
export const NUM_CHARCOAL_DASH = 5
export const NUM_ARROW = 6
export const NUM_LIME_NAVY = 8

const BULLET_STYLES = new Map<string, readonly [string, number]>([
  ['green-square', ['BwBBulletGreenSquare', NUM_GREEN_SQUARE]],
  ['navy-circle', ['BwBBulletNavyCircle', NUM_NAVY_CIRCLE]],
  ['green-circle', ['BwBBulletGreenCircle', NUM_GREEN_CIRCLE]],
  ['charcoal-dash', ['BwBBulletCharcoalDash', NUM_CHARCOAL_DASH]],
  ['arrow', ['BwBBulletArrow', NUM_ARROW]],
  ['lime-navy', ['BwBBulletLimeNavy', NUM_LIME_NAVY]],
])
const DEFAULT_BULLET: readonly [string, number] = ['BwBBulletGreenSquare', NUM_GREEN_SQUARE]

const ICON_CM: Readonly<Record<number, number>> = { 1: 0.7, 2: 0.55, 3: 0.5 }
const ICON_PT: Readonly<Record<number, number>> = { 1: 18, 2: 14, 3: 12 }

/** Python's int() of a JSON value: numbers are truncated, numeric strings parsed. */
export function toInt(v: unknown, fallback: number): number {
  const n = Math.trunc(Number(v))
  return Number.isFinite(n) ? n : fallback
}

/** Doc.p: a paragraph with the inline mark-up. */
export function paragraph(doc: Doc, text: unknown, style: string | null = null, kw: ParaProps = {}): void {
  const k = { ...doc.flags(), ...kw }
  doc.add(para(doc.runs(text), { style, ...k }))
}

export function heading(doc: Doc, text: string, levelIn: unknown = 1, numberedIn: unknown = true, inToc: unknown = true): void {
  const level = Math.max(1, Math.min(toInt(levelIn, 1), 6))
  const numbered = level > 3 ? false : Boolean(numberedIn)
  const style = numbered || level > 3 ? `Heading${level}` : `BwBHeading${level}Unnumbered`
  let label = ''
  if (numbered) {
    const c = doc.counters
    c[level - 1] = (c[level - 1] ?? 0) + 1
    for (let i = level; i < 3; i++) c[i] = 0
    const [a, b, d] = [c[0] ?? 0, c[1] ?? 0, c[2] ?? 0]
    label = level === 1 ? `${a}.` : level === 2 ? `${a}.${b}` : `${a}.${b}.${d}`
    if (level > 1 && a === 0) doc.warn(`Numbered level-${level} heading '${text}' appears before any numbered level-1 heading.`)
  }
  doc.bm += 1
  const bm = `_Toc${100000000 + doc.bm}`
  const flags = doc.flags()
  if (level === 1 && doc.h1PageBreak && !doc.sectionStart) flags.page_break_before = true
  const plain = pyStrip(fieldsAsText(stripIconTokens(text).replace(/\*\*|\^\[[^\n]*?\]/g, '')))
  doc.headings.push({ level, label, text: plain, bm, in_toc: Boolean(inToc) && level <= 6 })
  const head = `<w:p>${ppr({ style, ...flags })}`
  const runs = doc.runs(text, { icon_cm: ICON_CM[level] ?? 0.45, icon_pt: ICON_PT[level] ?? 11 })
  // insertOoxml would clash with the document's own _Toc bookmarks, so inserted headings carry none
  if (doc.fragment) doc.add(`${head}${runs}</w:p>`)
  else doc.add(`${head}<w:bookmarkStart w:id="${doc.bm}" w:name="${bm}"/>${runs}<w:bookmarkEnd w:id="${doc.bm}"/></w:p>`)
}

function itemParts(it: ListItem): [string, number] {
  return typeof it === 'string' ? [it, 0] : [it.text ?? '', toInt(it.level ?? 0, 0)]
}

export function bullets(doc: Doc, items: readonly ListItem[], style: string = 'green-square'): void {
  const [sid, numid] = BULLET_STYLES.get(style) ?? DEFAULT_BULLET
  items.forEach((it, n) => {
    const [text, lvl] = itemParts(it)
    const kw: ParaProps = { ...doc.flags() }
    if (lvl) kw.num = [numid, Math.min(lvl, 2)]
    if (n === items.length - 1) kw.spacing = { after: 200 }
    doc.add(para(doc.runs(text), { style: sid, ...kw }))
  })
}

/** Navy numerals; every block starts a new list instance that restarts at 1. */
export function numbered(doc: Doc, items: readonly ListItem[]): void {
  const numid = doc.nextNum
  doc.nextNum += 1
  doc.extraNums.push(numid)
  items.forEach((it, n) => {
    const [text, lvl] = itemParts(it)
    const kw: ParaProps = { ...doc.flags() }
    kw.num = [numid, Math.min(lvl, 2)]
    if (n === items.length - 1) kw.spacing = { after: 200 }
    doc.add(para(doc.runs(text), { style: 'BwBNumberedNavy', ...kw }))
  })
}

/** "Table 3. Text. Source: ..." with the number in a SEQ field (Doc._caption). */
export function caption(doc: Doc, kind: 'Table' | 'Figure', textIn: unknown, source?: string | null): string {
  doc.seq[kind] += 1
  const n = doc.seq[kind]
  const style = kind === 'Table' ? 'BwBTableCaption' : 'BwBImageCaption'
  const text = String(textIn || '').replace(new RegExp(`^\\s*${kind}\\s+\\p{Nd}+\\s*[.:\\-]\\s*`, 'iu'), '')
  const label =
    run(`${kind} `, { rstyle: 'BwBCaptionLabel' }) +
    field(`SEQ ${kind} \\* ARABIC`, String(n), { rstyle: 'BwBCaptionLabel' }) +
    run('.', { rstyle: 'BwBCaptionLabel' })
  const body = text ? pyRstrip(' ' + text) : ''
  let xml = label + (body ? doc.runs(body) : '')
  if (source && !text.toLowerCase().includes('source:')) xml += doc.runs(` Source: ${source}`)
  return para(xml, { style })
}

export function pullQuote(doc: Doc, textIn: unknown, attribution?: string | null): void {
  let text = pyStrip(String(textIn))
  if (!text.startsWith('\u{201C}')) text = `\u{201C}${pyStripChars(text, '"')}\u{201D}`
  const kw = doc.flags()
  doc.add(para(doc.runs(text), { style: 'BwBPullQuote', ...kw }))
  if (attribution) doc.add(para(doc.runs(attribution), { style: 'BwBPullQuoteSource' }))
  else doc.add(para('', { style: 'BwBPullQuoteSource' }))
}

export function figure(doc: Doc, b: FigureBlock): void {
  doc.imgN += 1
  if (!(b.alt || b.caption)) doc.warn(`Figure ${doc.imgN} has no alt text or caption; add "alt" in the content file.`)
  const rid = doc.pkg.addMedia(`fig${doc.imgN}.${b.ext}`, b.data)
  const sized = b.widthPx > 0 && b.heightPx > 0
  const wpx = sized ? b.widthPx : 1600
  const hpx = sized ? b.heightPx : 900
  const maxW = TEXT_W * EMU_PER_TWIP
  let cx = b.width_cm ? Math.trunc(b.width_cm * 360000) : Math.min(maxW, Math.trunc((wpx / 96) * 914400))
  cx = Math.min(cx, maxW)
  let cy = Math.trunc((cx * hpx) / wpx)
  const maxH = 14 * 360000
  if (cy > maxH) {
    cy = maxH
    cx = Math.trunc((cy * wpx) / hpx)
  }
  const pic = inlinePicture(rid, cx, cy, doc.did(), b.alt || b.caption || 'Figure', b.alt || cpHead(b.caption || '', 120))
  const kw = doc.flags()
  doc.add(para(pic, { jc: 'center', spacing: { before: 80, after: 40, line: 240, lineRule: 'auto' }, keep_next: true, ...kw }))
  if (b.caption || b.source) doc.add(caption(doc, 'Figure', b.caption || '', b.source))
  else doc.add(spacer(120))
}

/** A writing prompt in a block-level rich text content control (tag BwBGuidance) that the author replaces. */
export function guidance(doc: Doc, text: string): void {
  const id = doc.nextControlId()
  const kw = doc.flags()
  const p = para(doc.runs(text), { style: 'BwBLead', ...kw })
  doc.add(
    `<w:sdt><w:sdtPr><w:alias w:val="Guidance"/><w:tag w:val="BwBGuidance"/><w:id w:val="${id}"/></w:sdtPr>` +
      `<w:sdtContent>${p}</w:sdtContent></w:sdt>`,
  )
}
