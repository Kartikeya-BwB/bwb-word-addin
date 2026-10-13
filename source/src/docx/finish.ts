/** Output: icon PNGs, footnotes and endnotes parts, numbered-list restarts and the saved package (Doc.save). */
import { TOC_SLOT, type Doc } from './builder.ts'
import { ICON_PX } from './iconPolicy.ts'
import { noteRuns } from './inline.ts'
import type { CoreProps } from './package.ts'
import { sect, tocXml } from './sections.ts'
import { ROOT_NS, XML_DECL } from './xml.ts'

// A 1 x 1 transparent PNG, embedded when an icon renderer cannot draw a PNG (Word still shows the SVG).
export const TRANSPARENT_PNG = new Uint8Array([
  0x89, 0x50, 0x4e, 0x47, 0x0d, 0x0a, 0x1a, 0x0a, 0x00, 0x00, 0x00, 0x0d, 0x49, 0x48, 0x44, 0x52, 0x00, 0x00, 0x00,
  0x01, 0x00, 0x00, 0x00, 0x01, 0x08, 0x06, 0x00, 0x00, 0x00, 0x1f, 0x15, 0xc4, 0x89, 0x00, 0x00, 0x00, 0x0d, 0x49,
  0x44, 0x41, 0x54, 0x78, 0xda, 0x63, 0x60, 0x60, 0x60, 0x60, 0x00, 0x00, 0x00, 0x05, 0x00, 0x01, 0x7a, 0xa8, 0x57,
  0x50, 0x00, 0x00, 0x00, 0x00, 0x49, 0x45, 0x4e, 0x44, 0xae, 0x42, 0x60, 0x82,
])

const SEPARATORS = (tag: 'footnote' | 'endnote'): string =>
  `<w:${tag} w:type="separator" w:id="-1"><w:p><w:pPr><w:spacing w:after="0" w:line="240" w:lineRule="auto"/></w:pPr>` +
  `<w:r><w:separator/></w:r></w:p></w:${tag}>` +
  `<w:${tag} w:type="continuationSeparator" w:id="0"><w:p><w:pPr><w:spacing w:after="0" w:line="240" w:lineRule="auto"/>` +
  `</w:pPr><w:r><w:continuationSeparator/></w:r></w:p></w:${tag}>`

export function footnotesXml(doc: Doc): string {
  let out = XML_DECL + `<w:footnotes ${ROOT_NS}>` + SEPARATORS('footnote')
  doc.footnotes.forEach((t, i) => {
    out +=
      `<w:footnote w:id="${i + 1}"><w:p><w:pPr><w:pStyle w:val="FootnoteText"/></w:pPr>` +
      '<w:r><w:rPr><w:rStyle w:val="FootnoteReference"/></w:rPr><w:footnoteRef/></w:r>' +
      `<w:r><w:t xml:space="preserve"> </w:t></w:r>${noteRuns(t)}</w:p></w:footnote>`
  })
  return out + '</w:footnotes>'
}

export function endnotesXml(): string {
  return XML_DECL + `<w:endnotes ${ROOT_NS}>` + SEPARATORS('endnote') + '</w:endnotes>'
}

/** The numbering part with one restart instance (numId 20 and up) for each numbered list in the body. */
export function numberingWithRestarts(numbering: string, extraNums: readonly number[]): string {
  const stripped = numbering.replace(/<w:num w:numId="(?:[2-9]\d)"[^>]*>[\s\S]*?<\/w:num>/g, '')
  const extra = extraNums
    .map(
      (n) =>
        `<w:num w:numId="${n}"><w:abstractNumId w:val="6"/><w:lvlOverride w:ilvl="0">` +
        '<w:startOverride w:val="1"/></w:lvlOverride></w:num>',
    )
    .join('')
  return stripped.split('</w:numbering>').join(extra + '</w:numbering>')
}

/** Draw the PNG fallback of every icon used, now that the body is complete. */
export async function renderIcons(doc: Doc): Promise<void> {
  const icons = doc.icons
  const jobs = [...doc.pendingIcons]
  doc.pendingIcons.clear()
  if (!icons || !jobs.length) return
  const pngs = await Promise.all(
    jobs.map(async ([, { name, colour }]) => {
      try {
        return await icons.png(name, colour, ICON_PX)
      } catch {
        return null
      }
    }),
  )
  jobs.forEach(([file, { name, colour }], i) => {
    const png = pngs[i]
    if (!png || !png.length) doc.warn(`Icon '${name}' (${colour}) could not be drawn as a PNG; Word shows the SVG version only.`)
    doc.pkg.set(`word/media/${file}`, png && png.length ? png : TRANSPARENT_PNG)
  })
}

/** Write document.xml, the notes parts and the list restarts, then save the package (Doc.save). */
export async function finishDocx(doc: Doc, core: CoreProps, customProperties?: Readonly<Record<string, string>>): Promise<Uint8Array> {
  await renderIcons(doc)
  const body = [...doc.body]
  const slot = body.indexOf(TOC_SLOT)
  if (slot >= 0) body[slot] = tocXml(doc)
  // the last section is always content
  const xml = XML_DECL + `<w:document ${ROOT_NS}><w:body>` + body.join('') + sect(doc, 'content') + '</w:body></w:document>'
  const pkg = doc.pkg
  pkg.set('word/document.xml', xml)
  pkg.set('word/footnotes.xml', footnotesXml(doc))
  pkg.set('word/endnotes.xml', endnotesXml())
  if (!pkg.relIdFor('footnotes.xml')) pkg.addRel('footnotes', 'footnotes.xml')
  if (!pkg.relIdFor('endnotes.xml')) pkg.addRel('endnotes', 'endnotes.xml')
  pkg.set('word/numbering.xml', numberingWithRestarts(pkg.text('word/numbering.xml'), doc.extraNums))
  return pkg.save({ core, now: doc.now, customProperties })
}
