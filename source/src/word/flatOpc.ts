// A minimal flat OPC package (the single-XML form Word's insertOoxml accepts) carrying a body fragment
// together with the BwB template's styles and list definitions, so Word imports any it lacks.

import stylesXml from '../docx/template/styles.xml?raw'
import numberingXml from '../docx/template/numbering.xml?raw'

const W_NS = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main'
const R_NS = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships'
const REL = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships'
const PKG_REL = 'application/vnd.openxmlformats-package.relationships+xml'
const WML = 'application/vnd.openxmlformats-officedocument.wordprocessingml.'

function stripDeclaration(xml: string): string {
  return xml.replace(/^\s*<\?xml[^>]*\?>\s*/, '')
}

function part(name: string, contentType: string, xml: string): string {
  return `<pkg:part pkg:name="${name}" pkg:contentType="${contentType}"><pkg:xmlData>${xml}</pkg:xmlData></pkg:part>`
}

/** Wraps WordprocessingML body content (paragraphs and tables) in a package for insertOoxml. */
export function bodyToFlatOpc(bodyXml: string): string {
  const rels =
    `<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">` +
    `<Relationship Id="rId1" Type="${REL}/officeDocument" Target="word/document.xml"/></Relationships>`
  const docRels =
    `<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">` +
    `<Relationship Id="rId1" Type="${REL}/styles" Target="styles.xml"/>` +
    `<Relationship Id="rId2" Type="${REL}/numbering" Target="numbering.xml"/></Relationships>`
  const doc = `<w:document xmlns:w="${W_NS}" xmlns:r="${R_NS}"><w:body>${bodyXml}</w:body></w:document>`
  return (
    `<pkg:package xmlns:pkg="http://schemas.microsoft.com/office/2006/xmlPackage">` +
    part('/_rels/.rels', PKG_REL, rels) +
    part('/word/_rels/document.xml.rels', PKG_REL, docRels) +
    part('/word/document.xml', `${WML}document.main+xml`, doc) +
    part('/word/styles.xml', `${WML}styles+xml`, stripDeclaration(stylesXml)) +
    part('/word/numbering.xml', `${WML}numbering+xml`, stripDeclaration(numberingXml)) +
    `</pkg:package>`
  )
}

export interface TemplateStyle {
  id: string
  type: 'paragraph' | 'character' | 'table'
  custom: boolean
}

/** Styles only the template's specimen pages use; never copied into a user's document. */
const TEMPLATE_ONLY = /^BwB(Specimen|SampleLabel|PageLabel|Swatch|Icon)/

export function listTemplateStyles(): TemplateStyle[] {
  const out: TemplateStyle[] = []
  for (const m of stylesXml.matchAll(/<w:style\b([^>]*)>/g)) {
    const attrs = m[1] ?? ''
    const id = /w:styleId="([^"]+)"/.exec(attrs)?.[1]
    const type = /w:type="([^"]+)"/.exec(attrs)?.[1]
    if (!id || (type !== 'paragraph' && type !== 'character' && type !== 'table')) continue
    out.push({ id, type, custom: /w:customStyle="1"/.test(attrs) })
  }
  return out
}

/**
 * Body content that references every custom BwB style once. Inserting it makes Word copy the style
 * definitions (and the list definitions they use) into the document; the caller then deletes it.
 */
export function styleCarrierBody(): string {
  const styles = listTemplateStyles().filter((s) => s.custom && !TEMPLATE_ONLY.test(s.id))
  const paras = styles
    .filter((s) => s.type === 'paragraph')
    .map((s) => `<w:p><w:pPr><w:pStyle w:val="${s.id}"/></w:pPr><w:r><w:t>.</w:t></w:r></w:p>`)
  const chars = styles
    .filter((s) => s.type === 'character')
    .map((s) => `<w:r><w:rPr><w:rStyle w:val="${s.id}"/></w:rPr><w:t>.</w:t></w:r>`)
  const tables = styles
    .filter((s) => s.type === 'table')
    .map(
      (s) =>
        `<w:tbl><w:tblPr><w:tblStyle w:val="${s.id}"/><w:tblW w:w="2000" w:type="dxa"/></w:tblPr>` +
        `<w:tblGrid><w:gridCol w:w="2000"/></w:tblGrid><w:tr><w:tc><w:tcPr><w:tcW w:w="2000" w:type="dxa"/></w:tcPr>` +
        `<w:p><w:r><w:t>.</w:t></w:r></w:p></w:tc></w:tr></w:tbl><w:p/>`,
    )
  return paras.join('') + `<w:p>${chars.join('')}</w:p>` + tables.join('')
}
