/** The committed template assets made by scripts/make-base-template.py. */
import { strFromU8, unzipSync } from 'fflate'
import { describe, expect, it } from 'vitest'
import { baseTemplate, coverPhoto, readBytes } from '../../test/helpers/fixtures.ts'
import { readImageInfo } from './imageInfo.ts'
import { parseRels } from './package.ts'
import numberingXml from './template/numbering.xml?raw'
import stylesXml from './template/styles.xml?raw'

const parts = unzipSync(baseTemplate())
const text = (name: string): string => strFromU8(parts[name] ?? new Uint8Array())

describe('bwb-base.docx', () => {
  it('holds the parts the builder needs and no sample media', () => {
    expect(Object.keys(parts).sort()).toEqual([
      '[Content_Types].xml', '_rels/.rels', 'docProps/app.xml', 'docProps/core.xml', 'word/_rels/document.xml.rels',
      'word/_rels/footer1.xml.rels', 'word/_rels/footer2.xml.rels', 'word/_rels/header1.xml.rels', 'word/document.xml',
      'word/endnotes.xml', 'word/fontTable.xml', 'word/footer1.xml', 'word/footer2.xml', 'word/footer3.xml',
      'word/footnotes.xml', 'word/header1.xml', 'word/header2.xml', 'word/media/header_logo.png',
      'word/media/logo_black.png', 'word/media/logo_white.png', 'word/media/navy.png', 'word/media/strip.png',
      'word/numbering.xml', 'word/settings.xml', 'word/styles.xml', 'word/theme/theme1.xml', 'word/webSettings.xml',
    ])
  })

  it('has a minimal body in the content section and asks Word to update fields', () => {
    const doc = text('word/document.xml')
    expect(doc).toContain('<w:body><w:p/><w:sectPr><w:headerReference w:type="default" r:id="rId7"/><w:footerReference w:type="default" r:id="rId10"/>')
    const rels = parseRels(text('word/_rels/document.xml.rels'))
    expect(rels.find((r) => r.id === 'rId7')?.target).toBe('header1.xml')
    expect(text('word/header1.xml')).toContain('<w:drawing>')
    expect(rels.find((r) => r.id === 'rId10')?.target).toBe('footer2.xml')
    expect(text('word/footer2.xml')).toContain(' PAGE ')
    expect(text('word/settings.xml')).toContain('<w:updateFields w:val="true"/>')
    expect(text('word/footnotes.xml')).not.toContain('w:id="1"')
  })

  it('points every relationship at a part that exists', () => {
    for (const name of Object.keys(parts).filter((n) => n.endsWith('.rels'))) {
      const dir = name.replace(/_rels\/[^/]+$/, '')
      for (const r of parseRels(text(name))) expect(parts[(dir + r.target).replace(/^\//, '')], `${name}: ${r.target}`).toBeDefined()
    }
  })

  it('ships the same styles and numbering as the insertOoxml path, with list definitions 1 to 8 only', () => {
    expect(stylesXml).toBe(text('word/styles.xml'))
    expect(numberingXml).toBe(text('word/numbering.xml'))
    expect([...numberingXml.matchAll(/<w:num w:numId="(\d+)"/g)].map((m) => m[1])).toEqual(['1', '2', '3', '4', '5', '6', '7', '8'])
  })
})

describe('brand assets', () => {
  it('keeps the cover photo about 1400 px wide and under 450 KB', () => {
    const cover = coverPhoto()
    expect(cover.length).toBeLessThan(450_000)
    expect(readImageInfo(cover)).toEqual({ ext: 'jpg', widthPx: 1400, heightPx: 1563 })
  })

  it('copies the logos from the template', () => {
    expect(readBytes('public', 'assets', 'brand', 'logo-white.png')).toEqual(parts['word/media/logo_white.png'])
    expect(readBytes('public', 'assets', 'brand', 'logo-black.png')).toEqual(parts['word/media/logo_black.png'])
    expect(readBytes('public', 'assets', 'brand', 'header-logo.png')).toEqual(parts['word/media/header_logo.png'])
  })
})
