/**
 * Doc: the state of one document (or one set of blocks) being assembled on a BwB template package. A port of the
 * state and shared helpers of builder.Doc; the blocks are in blocksText.ts, blocksTables.ts, blocksPanels.ts and
 * sections.ts, the inline mark-up in inline.ts and the output in finish.ts.
 *
 * Every helper keeps the Python evaluation order, because drawing ids, bookmark ids, relationship ids and footnote
 * numbers are handed out in that order and appear in the XML.
 */
import { footerDividerXml, headerBlankXml } from './drawing.ts'
import { inlineRuns, type RunOpts } from './inline.ts'
import type { Package } from './package.ts'
import type { IconColour, IconRenderer } from './types.ts'
import { cpHead, esc, replaceFirst, type ParaProps } from './xml.ts'

export interface HeadingRecord {
  level: number
  label: string
  text: string
  bm: string
  in_toc: boolean
}

export interface DocOptions {
  icons?: IconRenderer | undefined
  /** Brand images that are not in the base package, by media name (the cover photograph, cover_default.jpg). */
  extraMedia?: ReadonlyMap<string, Uint8Array>
  /** Blocks for insertOoxml: no sections, no footnotes, no heading bookmarks. */
  fragment?: boolean
  now?: Date
  /** Id of the first content control the add-in extensions create, minus one. */
  controlIdBase?: number
  /** Id of the first drawing (wp:docPr), minus one; the Python builder starts after 1000. */
  drawingIdBase?: number
}

interface HeaderFooterIds {
  hdrContent: string
  hdrBlank: string
  ftrCover: string
  ftrContent: string
  ftrDivider: string
}

const MONTHS = [
  'January', 'February', 'March', 'April', 'May', 'June',
  'July', 'August', 'September', 'October', 'November', 'December',
]

/** "5 October 2026" (builder.fmt_date). */
export function fmtDate(d: Date): string {
  return `${d.getDate()} ${MONTHS[d.getMonth()] ?? ''} ${d.getFullYear()}`
}

export function shadeXml(fill: string): string {
  return `<w:shd w:val="clear" w:color="auto" w:fill="${fill}"/>`
}

/** Placeholder in the body for the table of contents, which is written last, when every heading is known. */
export const TOC_SLOT = '@@TOC@@'

export class Doc {
  readonly pkg: Package
  readonly icons: IconRenderer | undefined
  readonly fragment: boolean
  readonly now: Date
  /** Page of each heading bookmark for the cached contents; empty, as the browser cannot lay pages out (Word updates the
   * contents when it opens the file). */
  readonly pageMap = new Map<string, number>()
  body: string[] = []
  headings: HeadingRecord[] = []
  counters = [0, 0, 0]
  footnotes: string[] = []
  bm = 0
  seq: Record<'Table' | 'Figure', number> = { Table: 0, Figure: 0 }
  nextNum = 20
  extraNums: number[] = []
  pb = false
  sectionStart = true
  imgN = 0
  warnings: string[] = []
  h1PageBreak = true
  tocLevels = 3
  tocTitle = 'Table of Contents'
  cellSz: number | null = null
  /** True while a navy panel is being written (icons then default to white). */
  navy = false
  iconsUsed = new Set<string>()
  /** PNG renderings still to make, by media file name. */
  readonly pendingIcons = new Map<string, { name: string; colour: IconColour }>()
  private didCounter: number
  private controlId: number
  private readonly notesSeen = new Set<string>()
  private readonly extraMedia: ReadonlyMap<string, Uint8Array>
  private readonly hf: HeaderFooterIds | null

  constructor(pkg: Package, o: DocOptions = {}) {
    this.pkg = pkg
    this.icons = o.icons
    this.fragment = o.fragment ?? false
    this.now = o.now ?? new Date()
    this.extraMedia = o.extraMedia ?? new Map()
    this.controlId = o.controlIdBase ?? 0
    this.didCounter = o.drawingIdBase ?? 1000
    this.hf = this.fragment ? null : this.classifyHeaderFooterParts()
  }

  /** Header and footer relationship ids for the three kinds of section. */
  headerFooter(): HeaderFooterIds {
    if (!this.hf) throw new Error('Sections are not available when building blocks for insertion.')
    return this.hf
  }

  private addPart(name: string, xml: string, rtype: string): string {
    this.pkg.set(`word/${name}`, xml)
    return this.pkg.addRel(rtype, name)
  }

  /** Find the five header and footer roles by content, not by file name, so a template re-saved in Word still works. */
  private classifyHeaderFooterParts(): HeaderFooterIds {
    let hdrLogo: string | null = null
    let hdrBlank: string | null = null
    let fCover: string | null = null
    let fContent: string | null = null
    let fDivider: string | null = null
    for (const r of this.pkg.rels().rels) {
      const kind = r.type.slice(r.type.lastIndexOf('/') + 1)
      if ((kind !== 'header' && kind !== 'footer') || !this.pkg.has(`word/${r.target}`)) continue
      const xml = this.pkg.text(`word/${r.target}`)
      const drawings = xml.split('<w:drawing').length - 1
      if (kind === 'header') {
        if (drawings && !hdrLogo) hdrLogo = r.id
        else if (!drawings && !hdrBlank) hdrBlank = r.id
      } else {
        const hasPage = /instrText[^>]*>\s*PAGE\b/.test(xml) || xml.includes('w:fldSimple w:instr=" PAGE')
        if (hasPage && !fContent) fContent = r.id
        else if (drawings >= 2 && !hasPage && !fCover) fCover = r.id
        else if (drawings === 0 && !hasPage && !fDivider) fDivider = r.id
      }
    }
    if (!hdrLogo || !fContent) throw new Error('The base template lacks the BwB content header (logo) or the Page X of Y footer.')
    hdrBlank ??= this.addPart('header_blank.xml', headerBlankXml(), 'header')
    fDivider ??= this.addPart('footer_blank.xml', footerDividerXml(), 'footer')
    fCover ??= fDivider
    return { hdrContent: hdrLogo, hdrBlank, ftrCover: fCover, ftrContent: fContent, ftrDivider: fDivider }
  }

  // ---- utils ----
  did(): number {
    this.didCounter += 1
    return this.didCounter
  }

  warn(msg: string): void {
    this.warnings.push(msg)
  }

  noteOnce(msg: string): void {
    if (!this.notesSeen.has(msg)) {
      this.notesSeen.add(msg)
      this.warn(msg)
    }
  }

  /** A content control id for the add-in extensions: positive and unique within the build. */
  nextControlId(): number {
    this.controlId += 1
    return this.controlId
  }

  /** Relationship id of a brand image; the image is added from the extra media or the package when missing. */
  mediaRid(name: string): string {
    const rid = this.pkg.relIdFor(`media/${name}`)
    if (rid) return rid
    const extra = this.extraMedia.get(name)
    if (extra) return this.pkg.addMedia(name, extra)
    if (this.pkg.has(`word/media/${name}`)) return this.pkg.addRel('image', `media/${name}`)
    if (name === 'cover_default.jpg') throw new Error('The cover photograph is missing: pass assets.cover (bwb-cover.jpg).')
    throw new Error(`Brand image missing: ${name}.`)
  }

  /** Register an icon PNG that is drawn after the body is built; returns its relationship id. */
  addIconPng(file: string, name: string, colour: IconColour): string {
    this.pendingIcons.set(file, { name, colour })
    return this.pkg.addMedia(file, new Uint8Array(0))
  }

  flags(): ParaProps {
    if (this.pb) {
      this.pb = false
      return { page_break_before: true }
    }
    return {}
  }

  add(xml: string): void {
    this.body.push(xml)
    this.sectionStart = false
  }

  pageBreak(): void {
    this.pb = true
  }

  runs(text: unknown, base: RunOpts = {}): string {
    return inlineRuns(this, text, base)
  }

  /** Keep the preceding paragraph on the same page as the block that follows it (lead-in text never strands). */
  keepPrev(): void {
    let last = this.body[this.body.length - 1]
    if (last === undefined) return
    if (!last.startsWith('<w:p>') || cpHead(last, 400).includes('<w:keepNext/>') || last.includes('<w:sectPr>')) return
    // a caption belongs to the block above it, not the one below
    if (/<w:pStyle w:val="(BwBTableCaption|BwBImageCaption|BwBSource)"/.test(cpHead(last, 200))) return
    // an empty spacer: do not chain stacked panels together
    if (!last.includes('<w:r>') && !last.includes('<w:r ')) return
    if (last.startsWith('<w:p><w:pPr>')) {
      if (cpHead(last, 200).includes('<w:pStyle ')) last = last.replace(/(<w:pStyle [^>]*\/>)/, '$1<w:keepNext/>')
      else last = replaceFirst(last, '<w:pPr>', '<w:pPr><w:keepNext/>')
    } else {
      last = replaceFirst(last, '<w:p>', '<w:p><w:pPr><w:keepNext/></w:pPr>')
    }
    this.body[this.body.length - 1] = last
  }

  blockTable(tbl: string): void {
    if (!this.pb) this.keepPrev()
    if (this.pb) {
      this.pb = false
      this.add('<w:p><w:pPr><w:pageBreakBefore/><w:spacing w:before="0" w:after="0" w:line="20" w:lineRule="exact"/></w:pPr></w:p>')
    }
    this.add(tbl)
  }

  tc(w: number, content: string, fill?: string | null, borders?: string | null, mar?: string | null, valign?: string | null): string {
    let pr = `<w:tcW w:w="${w}" w:type="dxa"/>`
    if (borders) pr += borders
    if (fill) pr += shadeXml(fill)
    if (mar) pr += mar
    if (valign) pr += `<w:vAlign w:val="${valign}"/>`
    return `<w:tc><w:tcPr>${pr}</w:tcPr>${content}</w:tc>`
  }

  tblWrap(style: string, widths: readonly number[], rowsXml: string, borders = '', look = '04A0', alt = ''): string {
    const grid = widths.map((w) => `<w:gridCol w:w="${w}"/>`).join('')
    const cap = alt ? `<w:tblCaption w:val="${esc(alt)}"/>` : ''
    const total = widths.reduce((a, b) => a + b, 0)
    return (
      `<w:tbl><w:tblPr><w:tblStyle w:val="${style}"/><w:tblW w:w="${total}" w:type="dxa"/>${borders}` +
      `<w:tblLayout w:type="fixed"/><w:tblLook w:val="${look}" w:firstRow="0" w:lastRow="0" w:firstColumn="0" ` +
      `w:lastColumn="0" w:noHBand="1" w:noVBand="1"/>${cap}</w:tblPr><w:tblGrid>${grid}</w:tblGrid>${rowsXml}</w:tbl>`
    )
  }
}
