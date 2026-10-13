/**
 * Differential test: every content file in test/fixtures/docx is built by the TypeScript builder and compared with the
 * output of the Python builder (scripts/gen-docx-fixtures.py), part by part.
 *
 * Normalisation, and why each rule is needed:
 *  1. The created and modified times in docProps/core.xml are masked before comparing: they are the moment of the
 *     build. The rest of core.xml (title, subject, keywords, description, creator) must match.
 *  2. Icon renderings (word/media/icon_*.png and icon_*.svg) are compared by name only. The Python builder recolours
 *     the skill's SVG files and draws them with cairosvg; the tests recolour lucide-static's files (the same 1.52.0
 *     shapes, written differently) and draw them with resvg, so the bytes differ while the XML that uses them must not.
 * Everything else must be byte-identical: all XML parts (document, numbering, footnotes, endnotes, relationships,
 * content types, styles, settings, headers, footers, app.xml), the brand images, the cover and the figures, the list of
 * parts, and the builder's warnings.
 */
import { createHash } from 'node:crypto'
import { strFromU8, unzipSync } from 'fflate'
import { describe, expect, it } from 'vitest'
import {
  baseTemplate,
  coverPhoto,
  expectedPart,
  firstDifference,
  FIXED_NOW,
  fixtureNames,
  loadFixture,
  loadManifest,
  SAVED_PARTS,
} from '../../test/helpers/fixtures.ts'
import { createTestIconRenderer } from '../../test/helpers/iconRenderer.ts'
import { buildDocx } from './compose.ts'

const base = baseTemplate()
const cover = coverPhoto()

function sha256(data: Uint8Array): string {
  return createHash('sha256').update(data).digest('hex')
}

function comparedByNameOnly(part: string): boolean {
  return part === 'docProps/core.xml' || /^word\/media\/icon_.+\.(png|svg)$/.test(part)
}

function maskBuildTime(part: string, xml: string): string {
  if (part !== 'docProps/core.xml') return xml
  return xml.replace(/(<dcterms:(created|modified) xsi:type="dcterms:W3CDTF">)[^<]*(<\/dcterms:\2>)/g, '$1(build time)$3')
}

describe.each(fixtureNames())('%s', (name) => {
  it('matches the Python builder part for part', { timeout: 60_000 }, async () => {
    const manifest = loadManifest(name)
    const { docx, warnings } = await buildDocx(loadFixture(name), { base, cover, icons: createTestIconRenderer(), now: FIXED_NOW })
    const parts = unzipSync(docx)
    for (const part of SAVED_PARTS) {
      const actual = parts[part]
      expect(actual, `${part} is missing`).toBeDefined()
      const got = maskBuildTime(part, strFromU8(actual ?? new Uint8Array()))
      expect(firstDifference(got, maskBuildTime(part, expectedPart(name, part))), part).toBeNull()
    }
    expect(Object.keys(parts).sort()).toEqual(Object.keys(manifest.parts).sort())
    for (const [part, { sha256: expected }] of Object.entries(manifest.parts)) {
      if (comparedByNameOnly(part)) continue
      expect(sha256(parts[part] ?? new Uint8Array()), part).toBe(expected)
    }
    expect(warnings).toEqual(manifest.warnings)
  })
})
