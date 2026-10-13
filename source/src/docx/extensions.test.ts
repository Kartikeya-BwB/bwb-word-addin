/** The add-in extensions to the Python builder, and how buildDocx copes with missing assets. */
import { strFromU8, unzipSync } from 'fflate'
import { describe, expect, it } from 'vitest'
import { baseTemplate, coverPhoto, FIXED_NOW } from '../../test/helpers/fixtures.ts'
import { createTestIconRenderer } from '../../test/helpers/iconRenderer.ts'
import { buildDocx } from './compose.ts'
import { TRANSPARENT_PNG } from './finish.ts'
import type { BuildAssets, Content } from './types.ts'

const base = baseTemplate()
const cover = coverPhoto()
const assets: BuildAssets = { base, cover, now: FIXED_NOW }

async function build(content: Content, extra: Partial<BuildAssets> = {}) {
  const { docx, warnings } = await buildDocx(content, { ...assets, ...extra })
  const files = unzipSync(docx)
  const part = (name: string): string => strFromU8(files[name] ?? new Uint8Array())
  return { files, part, doc: part('word/document.xml'), warnings }
}

describe('guidance and {field:...} placeholders', () => {
  it('builds content controls with unique positive ids, and keeps fields readable where controls cannot go', async () => {
    const { doc, part } = await build({
      meta: { title: 'Proposal for {field:Client}', meta_left: 'Prepared for {field:Client}', format: 'report', toc: true, disclaimer: false },
      blocks: [
        { type: 'heading', level: 1, text: 'Why {field:Client}' },
        { type: 'guidance', text: 'Explain what {field:Client} asked for.' },
        { type: 'paragraph', text: 'A note^[About {field:Client}.] and a table:' },
        { type: 'table', header: ['Client'], rows: [['{field:Client}']] },
      ],
    })
    expect(doc).toContain(
      '<w:sdt><w:sdtPr><w:alias w:val="Guidance"/><w:tag w:val="BwBGuidance"/><w:id w:val="4"/></w:sdtPr><w:sdtContent>' +
        '<w:p><w:pPr><w:pStyle w:val="BwBLead"/></w:pPr><w:r><w:t xml:space="preserve">Explain what </w:t></w:r>' +
        '<w:sdt><w:sdtPr><w:alias w:val="Client"/><w:tag w:val="BwBPlaceholder"/><w:id w:val="5"/></w:sdtPr>' +
        '<w:sdtContent><w:r><w:t xml:space="preserve">[Client]</w:t></w:r></w:sdtContent></w:sdt>',
    )
    const ids = [...doc.matchAll(/<w:tag w:val="BwB(?:Guidance|Placeholder)"\/><w:id w:val="(\d+)"\/>/g)].map((m) => Number(m[1]))
    expect(ids).toEqual([1, 2, 3, 4, 5, 6]) // cover title, cover meta, heading, guidance, its field, table cell
    // the cover title is set with the mark-up at 32 pt, so the field is a control there too
    expect(doc).toContain('<w:r><w:t xml:space="preserve">Proposal for </w:t></w:r><w:sdt><w:sdtPr><w:alias w:val="Client"/>')
    // the cached contents entry and the footnote show the label as text
    expect(doc).toContain('<w:t xml:space="preserve">Why [Client]</w:t>')
    expect(part('word/footnotes.xml')).toContain('<w:t xml:space="preserve">About [Client].</w:t>')
  })
})

describe('extension blocks', () => {
  it('renders a disclaimer block exactly like the closing disclaimer', async () => {
    const a = await build({ meta: { title: 'X', disclaimer: false }, blocks: [{ type: 'disclaimer' }] })
    const b = await build({ meta: { title: 'X' }, blocks: [] })
    expect(a.doc).toBe(b.doc)
    expect(a.doc).toContain('<w:p><w:pPr><w:pStyle w:val="BwBDisclaimerHeading"/><w:pageBreakBefore/></w:pPr>')
    expect(a.doc).toContain('BwB (FRN: 953637) is authorised and regulated by the UK Financial Conduct Authority.')
  })

  it('renders a contents block as an empty TOC field followed by a page break', async () => {
    const { doc } = await build({
      meta: { title: 'X', format: 'memo', disclaimer: false },
      blocks: [{ type: 'toc', levels: 3, title: 'Contents' }, { type: 'paragraph', text: 'Body' }],
    })
    expect(doc).toContain(
      '<w:sdt><w:sdtPr><w:id w:val="1"/><w:docPartObj><w:docPartGallery w:val="Table of Contents"/><w:docPartUnique/></w:docPartObj></w:sdtPr>' +
        '<w:sdtContent><w:p><w:pPr><w:pStyle w:val="BwBTOCHeading"/></w:pPr><w:r><w:t xml:space="preserve">Contents</w:t></w:r></w:p>' +
        '<w:p><w:pPr><w:pStyle w:val="TOC1"/></w:pPr><w:r><w:fldChar w:fldCharType="begin"/></w:r>' +
        '<w:r><w:instrText xml:space="preserve"> TOC \\o "1-3" \\h \\z \\u </w:instrText></w:r><w:r><w:fldChar w:fldCharType="separate"/></w:r>' +
        '<w:r><w:t xml:space="preserve">Right-click and choose Update Field to build the contents.</w:t></w:r></w:p>' +
        '<w:p><w:r><w:fldChar w:fldCharType="end"/></w:r></w:p></w:sdtContent></w:sdt>' +
        '<w:p><w:pPr><w:pageBreakBefore/></w:pPr><w:r><w:t xml:space="preserve">Body</w:t></w:r></w:p>',
    )
  })

  it('numbers a caption block in sequence with table captions', async () => {
    const { doc } = await build({
      meta: { title: 'X', format: 'memo', disclaimer: false },
      blocks: [
        { type: 'table', header: ['A'], rows: [['1']], caption: 'First.' },
        { type: 'caption', kind: 'Table', text: 'Table 9: Second.', source: 'BwB' },
      ],
    })
    const numbers = [...doc.matchAll(/SEQ Table \\\* ARABIC[\s\S]*?<w:t>(\d+)<\/w:t>/g)].map((m) => m[1])
    expect(numbers).toEqual(['1', '2'])
    expect(doc).toContain('<w:t xml:space="preserve"> Second.</w:t></w:r><w:r><w:t xml:space="preserve"> Source: BwB</w:t>')
  })

  it('writes custom document properties only when asked', async () => {
    const withProps = await build({ meta: { title: 'X', custom_properties: { 'BwB document type': 'Report' } }, blocks: [] })
    expect(withProps.part('docProps/custom.xml')).toContain('name="BwB document type"><vt:lpwstr>Report</vt:lpwstr>')
    expect(withProps.part('_rels/.rels')).toContain('Target="docProps/custom.xml"')
    const without = await build({ meta: { title: 'X' }, blocks: [] })
    expect(without.files['docProps/custom.xml']).toBeUndefined()
    expect(without.part('_rels/.rels')).not.toContain('custom')
  })
})

describe('assets', () => {
  it('needs the cover photo for a report but not for a memo', async () => {
    await expect(buildDocx({ meta: { title: 'X' }, blocks: [] }, { base })).rejects.toThrow('The cover photograph is missing')
    const memo = await buildDocx({ meta: { title: 'X', format: 'memo' }, blocks: [] }, { base })
    expect(memo.docx.length).toBeGreaterThan(0)
  })

  it('uses a cover image from the content instead of the BwB photo', async () => {
    const png = new Uint8Array(TRANSPARENT_PNG)
    const { files } = await build({ meta: { title: 'X', cover_image: { data: png, ext: 'png', widthPx: 1, heightPx: 1 } }, blocks: [] }, { cover: new Uint8Array() })
    expect(files['word/media/cover_custom.png']).toEqual(png)
    expect(files['word/media/cover_default.jpg']).toBeUndefined()
  })

  it('skips icons without a renderer and warns once', async () => {
    const { doc, files, warnings } = await build({
      meta: { title: 'X', format: 'memo', disclaimer: false },
      blocks: [{ type: 'paragraph', text: '{icon:leaf} {icon:globe}' }, { type: 'callout', icon: 'leaf', title: 'T', text: 'x' }],
    })
    expect(warnings).toEqual(['Icons were skipped because no icon renderer was supplied.'])
    expect(doc).not.toContain('<w:drawing>')
    expect(Object.keys(files).some((n) => n.includes('icon_'))).toBe(false)
  })

  it('draws each icon and colour once, and falls back to a transparent PNG when drawing fails', async () => {
    const icons = createTestIconRenderer()
    await build({ meta: { title: 'X', format: 'memo' }, blocks: [{ type: 'paragraph', text: '{icon:leaf} {icon:leaf} {icon:leaf:green} {icon:globe}' }] }, { icons })
    expect(icons.pngCalls).toBe(3)
    const failing = createTestIconRenderer({ failPng: true })
    const { files, warnings } = await build({ meta: { title: 'X', format: 'memo' }, blocks: [{ type: 'paragraph', text: '{icon:leaf}' }] }, { icons: failing })
    expect(files['word/media/icon_leaf_navy.png']).toEqual(TRANSPARENT_PNG)
    expect(files['word/media/icon_leaf_navy.svg']?.length).toBeGreaterThan(100)
    expect(warnings).toEqual(["Icon 'leaf' (navy) could not be drawn as a PNG; Word shows the SVG version only."])
  })
})
