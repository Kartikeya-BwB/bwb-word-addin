import { strFromU8, unzipSync } from 'fflate'
import { describe, expect, it } from 'vitest'
import { coreXml, customXml, Package, parseRels, relsXml, RT } from './package.ts'

function pkgWithDocument(body: string, rels: string): Package {
  return new Package([
    ['word/document.xml', `<w:document><w:body>${body}</w:body></w:document>`],
    ['word/_rels/document.xml.rels', rels],
  ])
}

describe('relationships', () => {
  it('reads attributes in any order and quoting, unescaping targets', () => {
    const rels = parseRels(
      '<?xml version="1.0"?><Relationships xmlns="x"><Relationship Target="https://a.org/?b=1&amp;c=2" TargetMode="External" ' +
        "Type='t/hyperlink' Id='rId9'/><Relationship Id=\"rId2\" Type=\"t/image\" Target=\"media/a.png\"></Relationship></Relationships>",
    )
    expect(rels).toEqual([
      { id: 'rId9', type: 't/hyperlink', target: 'https://a.org/?b=1&c=2', mode: 'External' },
      { id: 'rId2', type: 't/image', target: 'media/a.png', mode: null },
    ])
  })

  it('writes them in the Python format', () => {
    expect(relsXml([{ id: 'rId1', type: 'T', target: 'a&b', mode: 'External' }])).toBe(
      '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">' +
        '<Relationship Id="rId1" Type="T" Target="a&amp;b" TargetMode="External"/></Relationships>',
    )
  })

  it('reuses a relationship to the same target and numbers new ones from rId1000', () => {
    const pkg = pkgWithDocument('', relsXml([{ id: 'rId1000', type: RT + 'styles', target: 'styles.xml', mode: null }]))
    expect(pkg.addRel('image', 'media/a.png')).toBe('rId1001')
    expect(pkg.addRel('image', 'media/a.png')).toBe('rId1001')
    expect(pkg.addRel('hyperlink', 'https://x.org', true)).toBe('rId1002')
    expect(pkg.relIdFor('https://x.org')).toBe('rId1002')
    expect(pkg.rels().rels.at(-1)).toEqual({ id: 'rId1002', type: RT + 'hyperlink', target: 'https://x.org', mode: 'External' })
    expect(pkg.addMedia('b.png', new Uint8Array([1]))).toBe('rId1003')
    expect(pkg.has('word/media/b.png')).toBe(true)
  })
})

describe('prune', () => {
  it('drops unused image and hyperlink relationships and media nothing points to', () => {
    const pkg = pkgWithDocument(
      '<a r:embed="rId1"/><w:hyperlink r:id="rId3"/>',
      relsXml([
        { id: 'rId1', type: RT + 'image', target: 'media/used.png', mode: null },
        { id: 'rId2', type: RT + 'image', target: 'media/unused.png', mode: null },
        { id: 'rId3', type: RT + 'hyperlink', target: 'https://a.org', mode: 'External' },
        { id: 'rId4', type: RT + 'hyperlink', target: 'https://b.org', mode: 'External' },
        { id: 'rId5', type: RT + 'header', target: 'header1.xml', mode: null },
      ]),
    )
    for (const n of ['used.png', 'unused.png', 'header.png', 'orphan.png']) pkg.set(`word/media/${n}`, new Uint8Array([1]))
    pkg.set('word/_rels/header1.xml.rels', relsXml([{ id: 'rId1', type: RT + 'image', target: '../word/media/header.png', mode: null }]))
    pkg.prune()
    expect(pkg.rels().rels.map((r) => r.id)).toEqual(['rId1', 'rId3', 'rId5'])
    expect(pkg.names().filter((n) => n.startsWith('word/media/')).sort()).toEqual(['word/media/header.png', 'word/media/used.png'])
  })
})

describe('content types and saving', () => {
  it('lists defaults by extension and overrides for known parts', () => {
    const pkg = new Package([
      ['word/document.xml', '<w:document/>'],
      ['word/header2.xml', '<w:hdr/>'],
      ['word/footer10.xml', '<w:ftr/>'],
      ['word/header_blank.xml', '<w:hdr/>'],
      ['word/media/a.PNG', new Uint8Array([1])],
      ['word/media/b.webp', new Uint8Array([1])],
      ['docProps/core.xml', '<x/>'],
    ])
    expect(pkg.contentTypes()).toBe(
      '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">' +
        '<Default Extension="png" ContentType="image/png"/><Default Extension="xml" ContentType="application/xml"/>' +
        '<Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>' +
        '<Override PartName="/docProps/core.xml" ContentType="application/vnd.openxmlformats-package.core-properties+xml"/>' +
        '<Override PartName="/word/footer10.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.footer+xml"/>' +
        '<Override PartName="/word/header2.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.header+xml"/>' +
        '</Types>',
    )
  })

  it('writes the zip in the Python order, drops stale parts and adds properties', () => {
    const pkg = pkgWithDocument('', relsXml([]))
    pkg.set('docProps/custom.xml', '<old/>')
    pkg.set('customXml/item1.xml', '<old/>')
    const now = new Date(Date.UTC(2026, 9, 5, 9, 53, 41))
    const bytes = pkg.save({ core: { title: 'T & C', subject: '', keywords: '', description: '' }, now })
    const files = unzipSync(bytes)
    expect(Object.keys(files).slice(0, 2)).toEqual(['[Content_Types].xml', '_rels/.rels'])
    expect(Object.keys(files)).not.toContain('docProps/custom.xml')
    expect(Object.keys(files)).not.toContain('customXml/item1.xml')
    const core = strFromU8(files['docProps/core.xml'] ?? new Uint8Array())
    expect(core).toBe(coreXml({ title: 'T & C', subject: '', keywords: '', description: '' }, now))
    expect(core).toContain('<dc:title>T &amp; C</dc:title>')
    expect(core).toContain('<dcterms:created xsi:type="dcterms:W3CDTF">2026-10-05T09:53:41Z</dcterms:created>')
    const reopened = Package.open(bytes)
    expect(reopened.text('word/document.xml')).toBe('<w:document><w:body></w:body></w:document>')
  })

  it('writes Office custom properties with their relationship and content type', () => {
    const pkg = pkgWithDocument('', relsXml([]))
    const files = unzipSync(pkg.save({ customProperties: { 'BwB document type': 'Report', Client: 'A & B' } }))
    const custom = strFromU8(files['docProps/custom.xml'] ?? new Uint8Array())
    expect(custom).toBe(customXml({ 'BwB document type': 'Report', Client: 'A & B' }))
    expect(custom).toContain(
      '<property fmtid="{D5CDD505-2E9C-101B-9397-08002B2CF9AE}" pid="2" name="BwB document type"><vt:lpwstr>Report</vt:lpwstr></property>' +
        '<property fmtid="{D5CDD505-2E9C-101B-9397-08002B2CF9AE}" pid="3" name="Client"><vt:lpwstr>A &amp; B</vt:lpwstr></property>',
    )
    expect(strFromU8(files['_rels/.rels'] ?? new Uint8Array())).toContain(
      '<Relationship Id="rId4" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/custom-properties" Target="docProps/custom.xml"/>',
    )
    expect(strFromU8(files['[Content_Types].xml'] ?? new Uint8Array())).toContain(
      '<Override PartName="/docProps/custom.xml" ContentType="application/vnd.openxmlformats-officedocument.custom-properties+xml"/>',
    )
  })

  it('fails clearly on a missing part', () => {
    expect(() => new Package().text('word/document.xml')).toThrow('The package has no part word/document.xml.')
  })
})
