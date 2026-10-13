import { describe, expect, it } from 'vitest'
import { readFlatOpc, type FlatPart } from '../../test/helpers/flatToDocx.ts'
import { createTestIconRenderer } from '../../test/helpers/iconRenderer.ts'
import { everyBlock } from '../../test/helpers/sampleBlocks.ts'
import { base64, buildBlockOoxml } from './flatOpc.ts'
import { parseRels } from './package.ts'
import type { Block } from './types.ts'

function partXml(parts: Map<string, FlatPart>, name: string): string {
  const xml = parts.get(name)?.xml
  if (xml === undefined) throw new Error(`${name} is missing`)
  return xml
}

function body(parts: Map<string, FlatPart>): string {
  return /<w:body>([\s\S]*)<\/w:body>/.exec(partXml(parts, '/word/document.xml'))?.[1] ?? ''
}

describe('buildBlockOoxml', async () => {
  const icons = createTestIconRenderer()
  const { ooxml, warnings } = await buildBlockOoxml(everyBlock(), { icons, seqStart: { Table: 4, Figure: 7 } })
  const parts = readFlatOpc(ooxml)
  const doc = partXml(parts, '/word/document.xml')
  const rels = parseRels(partXml(parts, '/word/_rels/document.xml.rels'))

  it('is a well-formed flat OPC package with the parts insertOoxml needs', () => {
    expect(ooxml.startsWith('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n<?mso-application progid="Word.Document"?><pkg:package')).toBe(true)
    for (const name of ['/_rels/.rels', '/word/document.xml', '/word/_rels/document.xml.rels', '/word/styles.xml', '/word/numbering.xml']) {
      expect(parts.has(name), name).toBe(true)
    }
    expect(parts.get('/word/document.xml')?.contentType).toBe('application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml')
    expect(partXml(parts, '/_rels/.rels')).toContain('Target="word/document.xml"')
    expect(doc).not.toContain('<?xml')
  })

  it('points every relationship at a part in the package, hyperlinks outside it', () => {
    for (const r of rels) {
      if (r.mode === 'External') expect(r.type.endsWith('/hyperlink')).toBe(true)
      else expect(parts.has(`/word/${r.target}`), r.target).toBe(true)
    }
    expect(rels.find((r) => r.target === 'https://bwb.earth/?a=1&b=2')?.mode).toBe('External')
    for (const m of doc.matchAll(/r:(?:embed|id)="([^"]+)"/g)) expect(rels.some((r) => r.id === m[1]), m[1]).toBe(true)
  })

  it('uses only styles and lists that the package defines', () => {
    const styles = partXml(parts, '/word/styles.xml')
    const ids = new Set([...styles.matchAll(/w:styleId="([^"]+)"/g)].map((m) => m[1]))
    for (const m of doc.matchAll(/<w:(?:pStyle|rStyle|tblStyle) w:val="([^"]+)"\/>/g)) expect(ids.has(m[1]), m[1]).toBe(true)
    const numbering = partXml(parts, '/word/numbering.xml')
    const nums = new Set([...numbering.matchAll(/<w:num w:numId="(\d+)"/g)].map((m) => m[1]))
    for (const m of doc.matchAll(/<w:numId w:val="(\d+)"\/>/g)) expect(nums.has(m[1]), m[1]).toBe(true)
    expect(numbering).toContain('<w:num w:numId="20"><w:abstractNumId w:val="6"/><w:lvlOverride w:ilvl="0"><w:startOverride w:val="1"/></w:lvlOverride></w:num>')
    expect(numbering).toContain('<w:num w:numId="21">')
  })

  it('carries each icon as a PNG with its SVG, drawn once per name and colour', async () => {
    const media = [...parts.values()].filter((p) => p.name.startsWith('/word/media/'))
    const names = media.map((p) => p.name).sort()
    expect(names).toContain('/word/media/icon_leaf_navy.png')
    expect(names).toContain('/word/media/icon_leaf_navy.svg')
    expect(names).toContain('/word/media/icon_leaf_green.png')
    expect(names).toContain('/word/media/fig1.png')
    for (const p of media) {
      expect(p.data?.length, p.name).toBeGreaterThan(0)
      if (p.name.endsWith('.png')) expect(p.contentType).toBe('image/png')
      if (p.name.endsWith('.svg')) expect(p.contentType).toBe('image/svg+xml')
    }
    const png = parts.get('/word/media/icon_leaf_green.png')?.data
    expect(png).toEqual(await icons.png('leaf', 'green', 256))
    expect(icons.pngCalls).toBe(names.filter((n) => n.includes('/icon_') && n.endsWith('.png')).length + 1)
    expect(doc).toContain('<asvg:svgBlip xmlns:asvg="http://schemas.microsoft.com/office/drawing/2016/SVG/main"')
  })

  it('wraps guidance and {field:...} in rich text content controls with unique positive ids', () => {
    expect(doc).toMatch(/<w:sdt><w:sdtPr><w:alias w:val="Guidance"\/><w:tag w:val="BwBGuidance"\/><w:id w:val="\d+"\/><\/w:sdtPr><w:sdtContent><w:p><w:pPr><w:pStyle w:val="BwBLead"\/><\/w:pPr>/)
    const fields = [...doc.matchAll(/<w:sdt><w:sdtPr><w:alias w:val="Client name"\/><w:tag w:val="BwBPlaceholder"\/><w:id w:val="(\d+)"\/><\/w:sdtPr><w:sdtContent>(.*?)<\/w:sdtContent><\/w:sdt>/g)]
    expect(fields).toHaveLength(2)
    expect(fields[0]?.[2]).toBe('<w:r><w:t xml:space="preserve">[Client name]</w:t></w:r>')
    const ids = [...doc.matchAll(/<w:sdtPr>(?:<w:alias [^>]*>)?(?:<w:tag [^>]*>)?<w:id w:val="(-?\d+)"\/>/g)].map((m) => Number(m[1]))
    expect(ids.length).toBe(4)
    expect(new Set(ids).size).toBe(ids.length)
    for (const id of ids) expect(Number.isInteger(id) && id > 0 && id < 2 ** 31).toBe(true)
  })

  it('numbers drawings uniquely, from a random start so separate insertions rarely clash', async () => {
    const ids = [...doc.matchAll(/<wp:docPr id="(\d+)"/g)].map((m) => Number(m[1]))
    expect(ids.length).toBeGreaterThan(10)
    expect(new Set(ids).size).toBe(ids.length)
    for (const id of ids) expect(id > 1_000_000 && id < 2 ** 31).toBe(true)
    const again = readFlatOpc((await buildBlockOoxml(everyBlock(), { icons })).ooxml)
    const other = [...partXml(again, '/word/document.xml').matchAll(/<wp:docPr id="(\d+)"/g)].map((m) => Number(m[1]))
    expect(other.some((id) => ids.includes(id))).toBe(false)
  })

  it('turns footnotes into bracketed text and says so', () => {
    expect(doc).not.toContain('footnoteReference')
    expect(doc).toContain(
      '<w:t xml:space="preserve"> text and a note</w:t></w:r><w:r><w:t xml:space="preserve"> [</w:t></w:r><w:r><w:t xml:space="preserve">With </w:t></w:r>' +
        '<w:r><w:rPr><w:b/><w:bCs/></w:rPr><w:t xml:space="preserve">bold</w:t></w:r><w:r><w:t xml:space="preserve"> and </w:t></w:r>' +
        '<w:r><w:rPr><w:i/><w:iCs/></w:rPr><w:t xml:space="preserve">italic</w:t></w:r><w:r><w:t xml:space="preserve"> text.</w:t></w:r>' +
        '<w:r><w:t xml:space="preserve">]</w:t></w:r><w:r><w:t xml:space="preserve">.</w:t></w:r>',
    )
    expect(warnings).toContain('Footnotes cannot be inserted into an open document; their text was put in square brackets instead.')
    expect(parts.has('/word/footnotes.xml')).toBe(false)
  })

  it('numbers captions from seqStart', () => {
    expect(doc).toContain('<w:instrText xml:space="preserve"> SEQ Table \\* ARABIC </w:instrText></w:r><w:r><w:rPr><w:rStyle w:val="BwBCaptionLabel"/></w:rPr><w:fldChar w:fldCharType="separate"/></w:r><w:r><w:rPr><w:rStyle w:val="BwBCaptionLabel"/></w:rPr><w:t>4</w:t>')
    expect([...doc.matchAll(/SEQ Figure[\s\S]*?<w:t>(\d+)<\/w:t>/g)].map((m) => m[1])).toEqual(['7', '8'])
  })

  it('has no sections, no heading bookmarks and a contents field with a placeholder', () => {
    expect(doc).not.toContain('<w:sectPr')
    expect(doc).not.toContain('<w:bookmarkStart')
    expect(doc).toContain('<w:docPartGallery w:val="Table of Contents"/>')
    expect(doc).toContain(' TOC \\o "1-3" \\h \\z \\u ')
    expect(doc).toContain('Right-click and choose Update Field to build the contents.')
    expect(doc).toMatch(/<w:p><w:pPr><w:pageBreakBefore\/><\/w:pPr><w:r><w:t xml:space="preserve">This paragraph follows the contents/)
  })

  it('ends a table with the builder spacer so Word does not merge it with the text after', () => {
    expect(body(parts).endsWith('</w:tbl><w:p><w:pPr><w:spacing w:before="0" w:after="0" w:line="120" w:lineRule="exact"/><w:rPr><w:sz w:val="2"/><w:szCs w:val="2"/></w:rPr></w:pPr></w:p>')).toBe(true)
    expect(doc).toContain('<w:tblCaption w:val="Disclaimer"/>')
  })
})

describe('buildBlockOoxml edge cases', () => {
  const run = async (blocks: Block[]) => {
    const { ooxml, warnings } = await buildBlockOoxml(blocks)
    return { body: body(readFlatOpc(ooxml)), warnings }
  }

  it('ends with an empty paragraph after text, so the last real paragraph keeps its style', async () => {
    const { body: b } = await run([{ type: 'lead', text: 'Lead' }])
    expect(b).toBe('<w:p><w:pPr><w:pStyle w:val="BwBLead"/></w:pPr><w:r><w:t xml:space="preserve">Lead</w:t></w:r></w:p><w:p/>')
  })

  it('keeps the panel spacer as the final empty paragraph', async () => {
    const { body: b } = await run([{ type: 'callout', text: 'x' }])
    expect(b.endsWith('</w:tbl><w:p><w:pPr><w:spacing w:before="0" w:after="0" w:line="160" w:lineRule="exact"/><w:rPr><w:sz w:val="2"/><w:szCs w:val="2"/></w:rPr></w:pPr></w:p>')).toBe(true)
  })

  it('writes a trailing page break as a break run', async () => {
    const { body: b } = await run([{ type: 'paragraph', text: 'x' }, { type: 'page_break' }])
    expect(b.endsWith('<w:p><w:r><w:br w:type="page"/></w:r></w:p><w:p/>')).toBe(true)
  })

  it('skips dividers and unknown blocks with a warning, and copes with no blocks', async () => {
    const r = await run([{ type: 'divider', title: 'x' }, { type: 'mystery' } as unknown as Block])
    expect(r.warnings).toEqual(['A section divider cannot be inserted into an open document; it was skipped.', "Unknown block type 'mystery' skipped."])
    expect(r.body).toBe('<w:p/>')
  })

  it('skips icons without a renderer, with one warning', async () => {
    const r = await run([{ type: 'paragraph', text: '{icon:leaf} and {icon:globe}' }])
    expect(r.warnings).toEqual(['Icons were skipped because no icon renderer was supplied.'])
    expect(r.body).not.toContain('<w:drawing>')
  })
})

describe('base64', () => {
  it('matches the platform encoder', () => {
    for (const n of [0, 1, 2, 3, 4, 5, 255, 1000]) {
      const d = new Uint8Array(n).map((_, i) => (i * 37 + n) & 255)
      expect(base64(d)).toBe(Buffer.from(d).toString('base64'))
    }
  })
})
