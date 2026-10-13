/** Sections and whole-page parts: section breaks, cover, dividers, memo title block, contents, disclaimer (builder.Doc). */
import { fmtDate, TOC_SLOT, type Doc } from './builder.ts'
import { anchoredPicture, inlinePicture, STRIP_H, STRIP_W, STRIP_Y_CONTENT } from './drawing.ts'
import { plainRuns } from './inline.ts'
import { wrap } from './textwrap.ts'
import type { PictureData } from './types.ts'
import {
  CHAR20,
  cpHead,
  EMU_PER_TWIP,
  esc,
  field,
  MARGIN,
  NAVY,
  OFFWHITE,
  PAGE_H,
  PAGE_H_EMU,
  PAGE_W,
  PAGE_W_EMU,
  para,
  pyStrip,
  replaceFirst,
  run,
  spacer,
  TEXT_W,
  type ParaProps,
} from './xml.ts'

export const DEFAULT_DISCLAIMER: readonly string[] = [
  'This document has been prepared by Bankers without Boundaries (\u{2018}BwB\u{2019}) for informational purposes only. It does not constitute investment advice or an offer, solicitation, or recommendation in relation to any transaction or investment.',
  'While the information contained in this document has been obtained from sources believed to be reliable, BwB makes no representation or warranty as to its accuracy or completeness and accepts no liability for any loss arising from reliance upon it. Any reference to specific organisations, products, or services does not constitute endorsement by BwB.',
  'Any forecasts, projections, or forward-looking statements contained in this document are based on assumptions and current expectations and are not guarantees of future performance.',
  'This document has been prepared without taking account of the objectives, financial situation, or needs of any particular person. Any person considering a transaction or investment should assess its suitability in light of their own circumstances, conduct their own due diligence, and seek independent professional advice before making any investment decisions. Nothing in this document constitutes legal, tax, accounting, or regulatory advice.',
  'The information and opinions expressed in this document are as of the date of writing and are subject to change without notice.',
  '**Confidentiality:** This document is intended solely for the designated recipients. It is confidential and must not be copied, quoted, discussed, or distributed to any other person, in whole or in part, without prior written approval from BwB.',
]
export const DISCLAIMER_NOTE = 'BwB (FRN: 953637) is authorised and regulated by the UK Financial Conduct Authority.'
export const TOC_PLACEHOLDER = 'Right-click and choose Update Field to build the contents.'

export type SectionKind = 'cover' | 'divider' | 'content'

export function sect(doc: Doc, kind: SectionKind): string {
  const hf = doc.headerFooter()
  const hdr = kind === 'content' ? hf.hdrContent : hf.hdrBlank
  const ftr = kind === 'cover' ? hf.ftrCover : kind === 'divider' ? hf.ftrDivider : hf.ftrContent
  let mar: string
  if (kind === 'cover') {
    mar = `<w:pgMar w:top="0" w:right="${MARGIN}" w:bottom="0" w:left="${MARGIN}" w:header="0" w:footer="0" w:gutter="0"/>`
  } else if (kind === 'divider') {
    mar = '<w:pgMar w:top="0" w:right="0" w:bottom="0" w:left="0" w:header="0" w:footer="0" w:gutter="0"/>'
  } else {
    mar =
      `<w:pgMar w:top="${MARGIN}" w:right="${MARGIN}" w:bottom="${MARGIN}" w:left="${MARGIN}" w:header="283" ` +
      'w:footer="778" w:gutter="0"/>'
  }
  return (
    `<w:sectPr><w:headerReference w:type="default" r:id="${hdr}"/><w:footerReference w:type="default" r:id="${ftr}"/>` +
    `<w:pgSz w:w="${PAGE_W}" w:h="${PAGE_H}" w:code="9"/>${mar}<w:cols w:space="720"/><w:docGrid w:linePitch="360"/></w:sectPr>`
  )
}

/** Close the section that was open (of the given kind) by attaching a sectPr to the last paragraph. */
export function endSection(doc: Doc, kind: SectionKind): void {
  const s = sect(doc, kind)
  const last = doc.body[doc.body.length - 1]
  if (last !== undefined && last.startsWith('<w:p>') && !last.includes('<w:sectPr>') && !cpHead(last, 200).includes('<w:fldChar')) {
    if (cpHead(last, 12).includes('<w:pPr>')) {
      doc.body[doc.body.length - 1] = last.includes('</w:pPr>') ? replaceFirst(last, '</w:pPr>', s + '</w:pPr>') : last
    } else {
      doc.body[doc.body.length - 1] = replaceFirst(last, '<w:p>', `<w:p><w:pPr>${s}</w:pPr>`)
    }
  } else {
    doc.body.push(`<w:p><w:pPr><w:spacing w:before="0" w:after="0" w:line="20" w:lineRule="exact"/>${s}</w:pPr></w:p>`)
  }
  doc.sectionStart = true
  doc.pb = false
}

export interface CoverOptions {
  tagline?: string | null | undefined
  metaLeft?: string | null | undefined
  metaRight?: string | null | undefined
  image?: PictureData | null | undefined
}

/**
 * Full cover page: photo on top, navy panel (a table, so the text can never fall outside it), title, tagline, lime
 * hairline and meta line. The cover is its own section: no header, page number hidden.
 */
export function cover(doc: Doc, titleIn: string, subtitleIn: string | null | undefined, o: CoverOptions = {}): void {
  let rid: string
  let crop: string
  if (o.image) {
    rid = doc.pkg.addMedia(`cover_custom.${o.image.ext}`, o.image.data)
    const t = PAGE_W_EMU / 7770495
    const a = o.image.widthPx / o.image.heightPx
    if (a < t) crop = `<a:srcRect b="${Math.trunc((1 - a / t) * 100000)}"/>`
    else {
      const e = Math.trunc(((1 - t / a) / 2) * 100000)
      crop = `<a:srcRect l="${e}" r="${e}"/>`
    }
  } else {
    rid = doc.mediaRid('cover_default.jpg')
    crop = '<a:srcRect b="8000"/>'
  }
  const photo = anchoredPicture(rid, PAGE_W_EMU, 7770495, doc.did(), 'page', 0, 'page', 0, {
    name: 'Cover image',
    descr: 'Cover photograph',
    behind: true,
    z: 251658240,
    crop,
  })
  const photoSdt =
    '<w:sdt><w:sdtPr><w:alias w:val="Cover image"/><w:tag w:val="BwBCoverImage"/>' +
    `<w:id w:val="100000001"/><w:picture/></w:sdtPr><w:sdtContent>${photo}</w:sdtContent></w:sdt>`
  const panelTopTw = Math.floor(6510137 / EMU_PER_TWIP)
  const panelHTw = Math.floor(3051958 / EMU_PER_TWIP)
  // fit the title: no more than two lines at the largest size that works (the width factor is conservative)
  const title = pyStrip(String(titleIn))
  const subtitle = pyStrip(String(subtitleIn || ''))
  let size = 32
  for (size of [32, 30, 28, 26, 24, 22]) {
    const cpl = Math.trunc(TEXT_W / 20 / (0.58 * size))
    const n = (wrap(title, cpl).length || 1) + (subtitle ? wrap(subtitle, cpl).length : 0)
    if (n <= 2) break
  }
  const line = Math.trunc(size * 22.5)
  doc.add(para(photoSdt, { spacing: { before: 0, after: 0, line: panelTopTw, lineRule: 'exact' } }))
  const mar =
    `<w:tcMar><w:top w:w="200" w:type="dxa"/><w:left w:w="${MARGIN}" w:type="dxa"/>` +
    `<w:bottom w:w="0" w:type="dxa"/><w:right w:w="${MARGIN}" w:type="dxa"/></w:tcMar>`
  const marB =
    `<w:tcMar><w:top w:w="0" w:type="dxa"/><w:left w:w="${MARGIN}" w:type="dxa"/>` +
    `<w:bottom w:w="200" w:type="dxa"/><w:right w:w="${MARGIN}" w:type="dxa"/></w:tcMar>`
  const titleKw: ParaProps = { style: 'Title', spacing: { before: 0, after: 40, line, lineRule: 'exact' } }
  // below 32 pt the Python builder sets the title as plain text with an explicit size, without the mark-up
  let up = para(size !== 32 ? plainRuns(doc, title, { sz: size }) : doc.runs(title), titleKw)
  if (subtitle) up += para(size !== 32 ? plainRuns(doc, subtitle, { sz: size }) : doc.runs(subtitle), titleKw)
  up += o.tagline ? para(doc.runs(o.tagline), { style: 'Subtitle' }) : ''
  const left = o.metaLeft ?? `Bankers without Boundaries  \u{B7}  ${fmtDate(doc.now)}`
  const low =
    para('', { style: 'BwBCoverRule' }) +
    para(plainRuns(doc, left) + (o.metaRight ? plainRuns(doc, '\t' + o.metaRight) : ''), { style: 'BwBCoverMeta' })
  const rowA = panelHTw - 1150
  const rows =
    `<w:tr><w:trPr><w:cantSplit/><w:trHeight w:val="${rowA}" w:hRule="atLeast"/></w:trPr>` +
    doc.tc(PAGE_W, up, NAVY, null, mar, 'center') +
    '</w:tr>' +
    '<w:tr><w:trPr><w:cantSplit/><w:trHeight w:val="1150" w:hRule="atLeast"/></w:trPr>' +
    doc.tc(PAGE_W, low, NAVY, null, marB, 'top') +
    '</w:tr>'
  const tbl =
    '<w:tbl><w:tblPr><w:tblStyle w:val="BwBPlain"/>' +
    `<w:tblW w:w="${PAGE_W}" w:type="dxa"/><w:tblInd w:w="${-MARGIN}" w:type="dxa"/><w:tblLayout w:type="fixed"/>` +
    '<w:tblLook w:val="0000" w:firstRow="0" w:lastRow="0" w:firstColumn="0" w:lastColumn="0" w:noHBand="1" w:noVBand="1"/>' +
    '<w:tblCaption w:val="Cover title panel"/></w:tblPr>' +
    `<w:tblGrid><w:gridCol w:w="${PAGE_W}"/></w:tblGrid>${rows}</w:tbl>`
  doc.add(tbl)
  endSection(doc, 'cover')
}

/** Full-page section divider (its own section, no header, colour strip drawn in the page). */
export function divider(doc: Doc, number: unknown, title: string, subtitle: string | null | undefined = '', variant = 'dark'): void {
  const dark = variant !== 'light'
  if (doc.body.length && !doc.sectionStart) endSection(doc, 'content')
  const sfx = dark ? '' : 'Light'
  let bg = ''
  if (dark) {
    bg = anchoredPicture(doc.mediaRid('navy.png'), PAGE_W_EMU, PAGE_H_EMU, doc.did(), 'page', 0, 'page', 0, {
      name: 'Divider background',
      descr: '',
      behind: true,
      z: 251658240,
    })
  }
  const top = 1100
  const strip = anchoredPicture(doc.mediaRid('strip.png'), STRIP_W, STRIP_H, doc.did(), 'page', 0, 'page', STRIP_Y_CONTENT, {
    name: 'BwB colour strip',
    descr: '',
    behind: false,
    halign: 'center',
    z: 251660288,
  })
  doc.add(para(bg + strip, { spacing: { before: 0, after: 0, line: top, lineRule: 'exact' } }))
  let logoH: number
  if (dark) {
    const logo = inlinePicture(doc.mediaRid('logo_white.png'), 1383665, 678815, doc.did(), 'BwB logo', 'Bankers without Boundaries logo')
    doc.add(para(logo, { jc: 'right', ind: { right: MARGIN }, spacing: { before: 0, after: 0, line: 1100, lineRule: 'exact' } }))
    logoH = 1100
  } else {
    const logo = inlinePicture(doc.mediaRid('logo_black.png'), 1900000, 968000, doc.did(), 'BwB logo', 'Bankers without Boundaries logo')
    doc.add(para(logo, { jc: 'left', ind: { left: MARGIN }, spacing: { before: 0, after: 0, line: 1600, lineRule: 'exact' } }))
    logoH = 1600
  }
  doc.add(spacer(Math.max(4600 - top - logoH, 200)))
  doc.add(para(doc.runs(String(number)), { style: `BwBSectionDividerNumber${sfx}` }))
  doc.add(para(doc.runs(title), { style: `BwBSectionDividerTitle${sfx}` }))
  if (subtitle) doc.add(para(doc.runs(subtitle), { style: `BwBSectionDividerSubtitle${sfx}` }))
  endSection(doc, 'divider')
}

/** Title block for documents without a cover (memos, notes, short briefs). */
export function titleBlock(doc: Doc, title: string, subtitle?: string | null, meta?: string | null): void {
  doc.add(para(doc.runs(String(title)), { style: 'BwBDocTitle', ...doc.flags() }))
  if (subtitle) doc.add(para(doc.runs(String(subtitle)), { style: 'BwBLead', spacing: { before: 0, after: 120 } }))
  const left = meta ?? `Bankers without Boundaries  \u{B7}  ${fmtDate(doc.now)}`
  doc.add(para(doc.runs(left), { style: 'BwBDocMeta' }))
  doc.sectionStart = true
}

/** Reserve the place of the table of contents; it is written when the document is saved. */
export function toc(doc: Doc, levels = 3, title = 'Table of Contents'): void {
  doc.tocLevels = levels
  doc.tocTitle = title
  doc.body.push(TOC_SLOT)
  doc.sectionStart = false
}

function tocFieldStart(levels: number): string {
  return (
    '<w:r><w:fldChar w:fldCharType="begin"/></w:r>' +
    `<w:r><w:instrText xml:space="preserve"> TOC \\o "1-${levels}" \\h \\z \\u </w:instrText></w:r>` +
    '<w:r><w:fldChar w:fldCharType="separate"/></w:r>'
  )
}

function tocSdt(id: number, inner: string): string {
  return (
    `<w:sdt><w:sdtPr><w:id w:val="${id}"/><w:docPartObj><w:docPartGallery w:val="Table of Contents"/>` +
    `<w:docPartUnique/></w:docPartObj></w:sdtPr><w:sdtContent>${inner}</w:sdtContent></w:sdt>`
  )
}

/** The contents: a TOC field whose cached entries link to the heading bookmarks (Doc._toc_xml). */
export function tocXml(doc: Doc): string {
  const ents = doc.headings.filter((h) => h.in_toc && h.level <= doc.tocLevels)
  const out = [para(plainRuns(doc, doc.tocTitle), { style: 'BwBTOCHeading' })]
  if (!ents.length) out.push(para(run('No headings found.'), { style: 'BwBTableText' }))
  ents.forEach((h, i) => {
    const page = String(doc.pageMap.get(h.bm) ?? 1)
    const begin = i === 0 ? tocFieldStart(doc.tocLevels) : ''
    let inner = ''
    if (h.label) {
      inner += `<w:r><w:rPr><w:noProof/></w:rPr><w:t>${esc(h.label)}</w:t></w:r><w:r><w:rPr><w:noProof/></w:rPr><w:tab/></w:r>`
    }
    inner += `<w:r><w:rPr><w:noProof/></w:rPr><w:t xml:space="preserve">${esc(h.text)}</w:t></w:r>`
    inner += '<w:r><w:rPr><w:noProof/></w:rPr><w:tab/></w:r>'
    inner += field(`PAGEREF ${h.bm} \\h`, page, { noproof: true })
    const link = `<w:hyperlink w:anchor="${h.bm}" w:history="1">${inner}</w:hyperlink>`
    out.push(para(begin + link, { style: `TOC${Math.min(h.level, 4)}` }))
  })
  out.push('<w:p><w:r><w:fldChar w:fldCharType="end"/></w:r></w:p>')
  return tocSdt(-1000001, out.join(''))
}

/**
 * Add-in extension: a contents field with no cached entries (Word builds them when fields are updated), then a page
 * break so the text starts on a new page.
 */
export function tocBlock(doc: Doc, levelsIn?: number, title = 'Table of Contents'): void {
  const levels = Math.max(1, Math.min(9, Math.trunc(Number(levelsIn ?? 2)) || 2))
  const id = doc.nextControlId()
  const head = para(plainRuns(doc, title), { style: 'BwBTOCHeading', ...doc.flags() })
  const entry = para(tocFieldStart(levels) + run(TOC_PLACEHOLDER), { style: 'TOC1' })
  doc.add(tocSdt(id, head + entry + '<w:p><w:r><w:fldChar w:fldCharType="end"/></w:r></w:p>'))
  doc.pageBreak()
}

/** The standard disclaimer on its own page, in an off-white box. */
export function disclaimer(
  doc: Doc,
  paragraphs?: readonly string[] | null,
  note: string | null | undefined = DISCLAIMER_NOTE,
  heading = 'Disclaimer',
): void {
  doc.pb = false
  doc.add(para(run(heading), { style: 'BwBDisclaimerHeading', page_break_before: true }))
  let inner = (paragraphs && paragraphs.length ? paragraphs : DEFAULT_DISCLAIMER)
    .map((t) => para(doc.runs(t), { style: 'BwBDisclaimerText' }))
    .join('')
  if (note) inner += para(doc.runs(note), { style: 'BwBDisclaimerNote' })
  const bord =
    '<w:tcBorders>' +
    ['top', 'left', 'bottom', 'right'].map((s) => `<w:${s} w:val="single" w:sz="4" w:space="0" w:color="${CHAR20}"/>`).join('') +
    '</w:tcBorders>'
  const mar =
    '<w:tcMar><w:top w:w="240" w:type="dxa"/><w:left w:w="280" w:type="dxa"/>' +
    '<w:bottom w:w="200" w:type="dxa"/><w:right w:w="280" w:type="dxa"/></w:tcMar>'
  const row = '<w:tr>' + doc.tc(TEXT_W, inner, OFFWHITE, bord, mar) + '</w:tr>'
  doc.add(doc.tblWrap('BwBPlain', [TEXT_W], row, '', '04A0', 'Disclaimer'))
}
