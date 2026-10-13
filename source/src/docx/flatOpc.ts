/**
 * BwB blocks as a flat OPC package for Word's range.insertOoxml: the body XML of the blocks with the BwB styles and list
 * definitions (so they come along when the destination document lacks them) and any media and hyperlinks.
 */
import numberingXml from './template/numbering.xml?raw'
import stylesXml from './template/styles.xml?raw'
import { Doc } from './builder.ts'
import { addBlock } from './compose.ts'
import { numberingWithRestarts, renderIcons } from './finish.ts'
import { CT_DEFAULTS, Package, REL_NS, RT } from './package.ts'
import type { Block, BlockOptions, BlockResult } from './types.ts'
import { ROOT_NS, spacer, XML_DECL } from './xml.ts'

const PKG_NS = 'http://schemas.microsoft.com/office/2006/xmlPackage'
const RELS_CT = 'application/vnd.openxmlformats-package.relationships+xml'
const WML = 'application/vnd.openxmlformats-officedocument.wordprocessingml.'
const B64 = 'ABCDEFGHIJKLMNOPQRSTUVWXYZabcdefghijklmnopqrstuvwxyz0123456789+/'

export function base64(d: Uint8Array): string {
  const out: string[] = []
  let i = 0
  for (; i + 2 < d.length; i += 3) {
    const n = ((d[i] ?? 0) << 16) | ((d[i + 1] ?? 0) << 8) | (d[i + 2] ?? 0)
    out.push(B64.charAt(n >> 18) + B64.charAt((n >> 12) & 63) + B64.charAt((n >> 6) & 63) + B64.charAt(n & 63))
  }
  const rest = d.length - i
  if (rest) {
    const n = ((d[i] ?? 0) << 16) | ((rest === 2 ? (d[i + 1] ?? 0) : 0) << 8)
    out.push(B64.charAt(n >> 18) + B64.charAt((n >> 12) & 63) + (rest === 2 ? B64.charAt((n >> 6) & 63) : '=') + '=')
  }
  return out.join('')
}

function stripDeclaration(xml: string): string {
  return xml.replace(/^\s*<\?xml[^>]*\?>\s*/, '')
}

/** The package as a pkg:package document, the form insertOoxml takes. */
export function toFlatOpc(pkg: Package): string {
  const parts: string[] = []
  const xmlPart = (name: string, ct: string, xml: string): void => {
    parts.push(`<pkg:part pkg:name="/${name}" pkg:contentType="${ct}"><pkg:xmlData>${stripDeclaration(xml)}</pkg:xmlData></pkg:part>`)
  }
  xmlPart(
    '_rels/.rels',
    RELS_CT,
    `<Relationships xmlns="${REL_NS}"><Relationship Id="rId1" Type="${RT}officeDocument" Target="word/document.xml"/></Relationships>`,
  )
  xmlPart('word/document.xml', WML + 'document.main+xml', pkg.text('word/document.xml'))
  xmlPart('word/_rels/document.xml.rels', RELS_CT, pkg.text('word/_rels/document.xml.rels'))
  xmlPart('word/styles.xml', WML + 'styles+xml', pkg.text('word/styles.xml'))
  xmlPart('word/numbering.xml', WML + 'numbering+xml', pkg.text('word/numbering.xml'))
  const media = pkg.names().filter((n) => n.startsWith('word/media/'))
  media.sort()
  for (const name of media) {
    const ext = name.slice(name.lastIndexOf('.') + 1).toLowerCase()
    const ct = (Object.hasOwn(CT_DEFAULTS, ext) ? CT_DEFAULTS[ext] : undefined) ?? 'application/octet-stream'
    parts.push(
      `<pkg:part pkg:name="/${name}" pkg:contentType="${ct}" pkg:compression="store">` +
        `<pkg:binaryData>${base64(pkg.get(name))}</pkg:binaryData></pkg:part>`,
    )
  }
  return XML_DECL + '<?mso-application progid="Word.Document"?>' + `<pkg:package xmlns:pkg="${PKG_NS}">${parts.join('')}</pkg:package>`
}

function firstNumber(v: number | undefined): number {
  return v !== undefined && Number.isFinite(v) && v >= 1 ? Math.trunc(v) : 1
}

function isEmptyParagraph(xml: string): boolean {
  return xml.startsWith('<w:p>') && !xml.includes('<w:r>') && !xml.includes('<w:r ') && !xml.includes('<w:sectPr>')
}

/**
 * BwB blocks as flat OPC XML for range.insertOoxml. Footnotes become bracketed text, dividers are skipped (both with a
 * warning), headings carry no bookmarks, and the content ends with an empty paragraph: Word merges the last inserted
 * paragraph into the paragraph at the insertion point, so that paragraph must carry no content of its own.
 */
export async function buildBlockOoxml(blocks: readonly Block[], options: BlockOptions = {}): Promise<BlockResult> {
  const pkg = new Package([
    ['word/styles.xml', stylesXml],
    ['word/numbering.xml', numberingXml],
  ])
  pkg.setRels('word/document.xml', [
    ['rId1', 'styles', 'styles.xml'],
    ['rId2', 'numbering', 'numbering.xml'],
  ])
  // content control and drawing ids start at a random point so that blocks inserted one after another, or into a
  // document built by buildDocx, rarely share an id with what is already there
  const controlIdBase = 1_000_000 + Math.floor(Math.random() * 1_000_000_000)
  const drawingIdBase = 1_000_000 + Math.floor(Math.random() * 1_000_000_000)
  const doc = new Doc(pkg, { icons: options.icons, fragment: true, controlIdBase, drawingIdBase })
  doc.h1PageBreak = false
  doc.seq.Table = firstNumber(options.seqStart?.Table) - 1
  doc.seq.Figure = firstNumber(options.seqStart?.Figure) - 1
  for (const b of blocks) addBlock(doc, b)
  if (doc.pb) doc.add('<w:p><w:r><w:br w:type="page"/></w:r></w:p>')
  const last = doc.body[doc.body.length - 1]
  if (last !== undefined && last.startsWith('<w:tbl>')) doc.add(spacer(120))
  else if (last === undefined || !isEmptyParagraph(last)) doc.add('<w:p/>')
  await renderIcons(doc)
  pkg.set('word/document.xml', XML_DECL + `<w:document ${ROOT_NS}><w:body>${doc.body.join('')}</w:body></w:document>`)
  pkg.set('word/numbering.xml', numberingWithRestarts(numberingXml, doc.extraNums))
  pkg.prune()
  return { ooxml: toFlatOpc(pkg), warnings: [...doc.warnings] }
}
