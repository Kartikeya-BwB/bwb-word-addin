// The in-memory document behind the fake Office.js. Plain data, so tests can set up and inspect documents.

import { DOMParser } from '@xmldom/xmldom'

export interface FakeFont {
  name: string | null
  size: number | null
  color: string | null
  highlightColor: string | null
  bold: boolean | null
  italic: boolean | null
}

export interface FakeParagraph {
  text: string
  style: string
  styleBuiltIn: string
  isListItem: boolean
  tableNestingLevel: number
  font: FakeFont
  alignment?: string
  table?: FakeTable
}

export interface FakeTable {
  style: string
  rows: string[][]
  headerRowCount?: number
  /** styleFirstColumn, styleTotalRow, styleBandedRows and styleBandedColumns as last set. */
  options?: Record<string, boolean>
  /** Horizontal alignment set on cells, keyed "row,column". */
  alignments?: Record<string, string>
}

export interface FakeControl {
  tag: string
  title: string
  appearance?: string
  paragraphs: FakeParagraph[]
}

export interface FakePart {
  paragraphs: FakeParagraph[]
  controls: FakeControl[]
  pictures: number
}

export interface FakeSection {
  headers: Record<'Primary' | 'FirstPage' | 'EvenPages', FakePart>
  footers: Record<'Primary' | 'FirstPage' | 'EvenPages', FakePart>
}

export interface FakeModel {
  paragraphs: FakeParagraph[]
  tables: FakeTable[]
  pictures: Array<{ altTextTitle: string; altTextDescription: string }>
  controls: FakeControl[]
  sections: FakeSection[]
  styles: Set<string>
  properties: { title: string; subject: string; author: string; company: string }
  custom: Map<string, string>
  fields: Array<{ code: string; updates: number }>
  comments: number
  trackedChanges: number
  selection: FakeParagraph[]
  selectedText: string
  /** The table the cursor is in, if any. */
  selectionTable?: FakeTable
  inserts: Array<{ location: string; ooxml: string }>
  created: string[]
  selected: string[]
  footnotes: string[]
  settings: Map<string, unknown>
  url: string
  apiVersion: number
  styleUpdates: Array<{ name: string; font?: unknown; paragraph?: unknown }>
  /** Emulates Word leaving an empty paragraph after inserted content is deleted at the end. */
  ooxmlQuirk?: boolean
}

export const font = (over: Partial<FakeFont> = {}): FakeFont => ({
  name: 'Aptos', size: 10, color: '#000000', highlightColor: null, bold: false, italic: false, ...over,
})

export function para(text: string, style = 'Normal', over: Partial<FakeParagraph> = {}): FakeParagraph {
  const builtIn = /^heading ([1-9])$/i.exec(style)
  return {
    text,
    style,
    styleBuiltIn: builtIn ? `Heading${builtIn[1]}` : style === 'Normal' ? 'Normal' : 'Other',
    isListItem: false,
    tableNestingLevel: 0,
    font: font(),
    ...over,
  }
}

const emptyPart = (): FakePart => ({ paragraphs: [], controls: [], pictures: 0 })

export function section(headerText = '', footerText = '', headerPictures = 0): FakeSection {
  const header = emptyPart()
  if (headerText) header.paragraphs.push(para(headerText, 'header'))
  header.pictures = headerPictures
  const footer = emptyPart()
  if (footerText) footer.paragraphs.push(para(footerText, 'footer'))
  return {
    headers: { Primary: header, FirstPage: emptyPart(), EvenPages: emptyPart() },
    footers: { Primary: footer, FirstPage: emptyPart(), EvenPages: emptyPart() },
  }
}

export function createModel(paragraphs: FakeParagraph[] = [para('')], over: Partial<FakeModel> = {}): FakeModel {
  return {
    paragraphs,
    tables: [],
    pictures: [],
    controls: [],
    sections: [section('', 'Page 1 of 1', 1)],
    styles: new Set(['Normal', 'Heading 1', 'Heading 2', 'Heading 3', 'Heading 4', 'Caption']),
    properties: { title: '', subject: '', author: '', company: '' },
    custom: new Map(),
    fields: [],
    comments: 0,
    trackedChanges: 0,
    selection: paragraphs.slice(0, 1),
    selectedText: '',
    inserts: [],
    created: [],
    selected: [],
    footnotes: [],
    settings: new Map(),
    url: '',
    apiVersion: 1.6,
    styleUpdates: [],
    ...over,
  }
}

const W = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main'
const PKG = 'http://schemas.microsoft.com/office/2006/xmlPackage'

type XmlElement = NonNullable<ReturnType<DOMParser['parseFromString']>['documentElement']>

function childElements(el: XmlElement): XmlElement[] {
  const out: XmlElement[] = []
  for (let n = el.firstChild; n; n = n.nextSibling) if (n.nodeType === 1) out.push(n as XmlElement)
  return out
}

function textOf(el: XmlElement): string {
  const ts = el.getElementsByTagNameNS(W, 't')
  let s = ''
  for (let i = 0; i < ts.length; i += 1) s += ts.item(i)?.textContent ?? ''
  return s
}

function valOf(el: XmlElement | undefined, tag: string): string | null {
  if (!el) return null
  const found = el.getElementsByTagNameNS(W, tag).item(0)
  return found ? found.getAttributeNS(W, 'val') || found.getAttribute('w:val') : null
}

export interface IngestResult {
  paragraphs: FakeParagraph[]
  tables: FakeTable[]
  controls: FakeControl[]
  styleNames: string[]
}

/** Turns an insertOoxml package into model content, the way Word would lay it out. */
export function ingestOoxml(ooxml: string): IngestResult {
  const doc = new DOMParser().parseFromString(ooxml, 'application/xml')
  const names = new Map<string, string>()
  const parts = doc.getElementsByTagNameNS(PKG, 'part')
  let body: XmlElement | undefined
  for (let i = 0; i < parts.length; i += 1) {
    const part = parts.item(i)
    if (!part) continue
    const name = part.getAttributeNS(PKG, 'name') || part.getAttribute('pkg:name')
    if (name === '/word/styles.xml') {
      const styles = part.getElementsByTagNameNS(W, 'style')
      for (let j = 0; j < styles.length; j += 1) {
        const s = styles.item(j)
        const id = s?.getAttributeNS(W, 'styleId') || s?.getAttribute('w:styleId')
        const n = s ? valOf(s, 'name') : null
        if (id && n) names.set(id, n)
      }
    }
    if (name === '/word/document.xml') body = part.getElementsByTagNameNS(W, 'body').item(0) ?? undefined
  }
  const result: IngestResult = { paragraphs: [], tables: [], controls: [], styleNames: [] }
  if (!body) return result
  const styleName = (id: string | null): string => (id ? (names.get(id) ?? id) : 'Normal')
  const used = new Set<string>()
  const addParagraph = (p: XmlElement, nesting: number, table?: FakeTable): FakeParagraph => {
    const pPr = p.getElementsByTagNameNS(W, 'pPr').item(0) ?? undefined
    const id = pPr ? valOf(pPr, 'pStyle') : null
    if (id) used.add(styleName(id))
    const fp = para(textOf(p), styleName(id), { tableNestingLevel: nesting, ...(table ? { table } : {}) })
    result.paragraphs.push(fp)
    return fp
  }
  const walk = (el: XmlElement, nesting: number, table?: FakeTable): void => {
    for (const child of childElements(el)) {
      if (child.namespaceURI !== W) continue
      if (child.localName === 'p') addParagraph(child, nesting, table)
      else if (child.localName === 'tbl') {
        const tblPr = child.getElementsByTagNameNS(W, 'tblPr').item(0) ?? undefined
        const t: FakeTable = { style: styleName(valOf(tblPr, 'tblStyle')), rows: [] }
        used.add(t.style)
        for (const tr of childElements(child).filter((c) => c.localName === 'tr')) {
          const row: string[] = []
          for (const tc of childElements(tr).filter((c) => c.localName === 'tc')) {
            row.push(textOf(tc))
            walk(tc, nesting + 1, t)
          }
          t.rows.push(row)
        }
        result.tables.push(t)
      } else if (child.localName === 'sdt') {
        const sdtPr = child.getElementsByTagNameNS(W, 'sdtPr').item(0) ?? undefined
        const content = child.getElementsByTagNameNS(W, 'sdtContent').item(0)
        const before = result.paragraphs.length
        if (content) walk(content, nesting, table)
        result.controls.push({ tag: valOf(sdtPr, 'tag') ?? '', title: valOf(sdtPr, 'alias') ?? '', paragraphs: result.paragraphs.slice(before) })
      }
    }
  }
  walk(body, 0)
  result.styleNames = [...used]
  return result
}
