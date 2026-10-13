/**
 * Node-only: read the flat OPC package that buildBlockOoxml returns, and merge one or more of them into the base
 * template as a normal .docx (roughly what Word does on insertOoxml), so inserted blocks can be rendered and checked.
 */
import { DOMParser } from '@xmldom/xmldom'
import { Package, parseRels, RT, type Rel } from '../../src/docx/package.ts'

export const PKG_NS = 'http://schemas.microsoft.com/office/2006/xmlPackage'

export interface FlatPart {
  name: string
  contentType: string
  /** XML parts: the serialised content of pkg:xmlData. */
  xml?: string
  /** Binary parts: the decoded bytes of pkg:binaryData. */
  data?: Uint8Array
}

/** Parse the flat package with a real XML parser (fails on malformed XML) and return its parts by name. */
export function readFlatOpc(ooxml: string): Map<string, FlatPart> {
  const errors: string[] = []
  const doc = new DOMParser({ onError: (level, msg) => void (level !== 'warning' && errors.push(msg)) }).parseFromString(
    ooxml,
    'text/xml',
  )
  if (errors.length) throw new Error(`flat OPC is not well-formed: ${errors.join('; ')}`)
  const out = new Map<string, FlatPart>()
  const parts = doc.getElementsByTagNameNS(PKG_NS, 'part')
  for (let i = 0; i < parts.length; i++) {
    const p = parts.item(i)
    if (!p) continue
    const name = p.getAttributeNS(PKG_NS, 'name') ?? ''
    const contentType = p.getAttributeNS(PKG_NS, 'contentType') ?? ''
    const xmlData = p.getElementsByTagNameNS(PKG_NS, 'xmlData').item(0)
    const bin = p.getElementsByTagNameNS(PKG_NS, 'binaryData').item(0)
    const part: FlatPart = { name, contentType }
    if (xmlData) {
      let xml = ''
      for (let c = xmlData.firstChild; c; c = c.nextSibling) xml += c.toString()
      part.xml = xml
    }
    if (bin) part.data = new Uint8Array(Buffer.from(bin.textContent ?? '', 'base64'))
    out.set(name, part)
  }
  return out
}

function bodyOf(documentXml: string): string {
  const m = /<w:body>([\s\S]*)<\/w:body>/.exec(documentXml)
  return m?.[1] ?? ''
}

/**
 * A .docx made from the base template with the blocks of each flat package appended in order, the relationships
 * renumbered, the media copied and the numbered-list restart instances merged.
 */
export function flatOpcToDocx(base: Uint8Array, fragments: readonly string[]): Uint8Array {
  const pkg = Package.open(base)
  const baseDoc = pkg.text('word/document.xml')
  const sect = /<w:sectPr>[\s\S]*<\/w:sectPr>/.exec(baseDoc)?.[0] ?? ''
  const head = baseDoc.slice(0, baseDoc.indexOf('<w:body>'))
  const docRels = pkg.rels().rels
  let numbering = pkg.text('word/numbering.xml')
  let body = ''
  let nextId = 5000
  fragments.forEach((ooxml, fi) => {
    const parts = readFlatOpc(ooxml)
    let xml = bodyOf(parts.get('/word/document.xml')?.xml ?? '')
    const rels = parseRels(parts.get('/word/_rels/document.xml.rels')?.xml ?? '')
    for (const r of rels) {
      const kind = r.type.slice(r.type.lastIndexOf('/') + 1)
      if (kind !== 'image' && kind !== 'hyperlink') continue
      const id = `rId${nextId++}`
      let target = r.target
      if (kind === 'image') {
        target = `media/f${fi}_${r.target.slice('media/'.length)}`
        const media = parts.get(`/word/${r.target}`)
        if (media?.data) pkg.set(`word/${target}`, media.data)
      }
      docRels.push({ id, type: RT + kind, target, mode: r.mode } satisfies Rel)
      xml = xml.replaceAll(`r:embed="${r.id}"`, `r:embed="${id}"`).replaceAll(`r:id="${r.id}"`, `r:id="${id}"`)
    }
    const nums = /<w:num w:numId="(?:[2-9]\d)"[\s\S]*?<\/w:num>/g
    for (const m of (parts.get('/word/numbering.xml')?.xml ?? '').matchAll(nums)) {
      const shift = 20 + fi * 10
      const oldId = /w:numId="(\d+)"/.exec(m[0])?.[1] ?? '0'
      const newId = String(Number(oldId) + shift)
      numbering = numbering.replace('</w:numbering>', m[0].replace(`w:numId="${oldId}"`, `w:numId="${newId}"`) + '</w:numbering>')
      xml = xml.replaceAll(`<w:numId w:val="${oldId}"/>`, `<w:numId w:val="${newId}"/>`)
    }
    body += xml
  })
  pkg.writeRels('word/_rels/document.xml.rels', docRels)
  pkg.set('word/numbering.xml', numbering)
  pkg.set('word/document.xml', `${head}<w:body>${body}${sect}</w:body></w:document>`)
  return pkg.save({ core: { title: 'Inserted blocks', subject: '', keywords: '', description: '' } })
}
