/**
 * Builds the sample documents and a round trip of inserted blocks. With BWB_DOCX_OUT set to a folder, the files are
 * also written there for scripts/render-docx.py and the skill's lint_docx.py:
 *   BWB_DOCX_OUT=/tmp/out npx vitest run src/docx/samples.test.ts
 */
import { mkdirSync, writeFileSync } from 'node:fs'
import { join } from 'node:path'
import { unzipSync } from 'fflate'
import { describe, expect, it } from 'vitest'
import { baseTemplate, coverPhoto, FIXED_NOW, loadFixture } from '../../test/helpers/fixtures.ts'
import { flatOpcToDocx, readFlatOpc } from '../../test/helpers/flatToDocx.ts'
import { createTestIconRenderer } from '../../test/helpers/iconRenderer.ts'
import { everyBlock } from '../../test/helpers/sampleBlocks.ts'
import { buildDocx } from './compose.ts'
import { buildBlockOoxml } from './flatOpc.ts'
import type { Content } from './types.ts'

const out = process.env['BWB_DOCX_OUT']
const base = baseTemplate()
const icons = createTestIconRenderer()

function save(name: string, bytes: Uint8Array): void {
  if (!out) return
  mkdirSync(out, { recursive: true })
  writeFileSync(join(out, name), bytes)
}

describe('sample documents', () => {
  it.each(['report-sample', 'memo-sample', 'icons-sample', 'panels', 'tables'])('builds %s', { timeout: 60_000 }, async (name) => {
    const { docx } = await buildDocx(loadFixture(name), { base, cover: coverPhoto(), icons, now: FIXED_NOW })
    expect(Object.keys(unzipSync(docx))).toContain('word/document.xml')
    save(`ts-${name}.docx`, docx)
  })

  it('builds a new document with writing prompts and placeholders', { timeout: 60_000 }, async () => {
    const content: Content = {
      meta: { title: 'Proposal for {field:Client}', tagline: 'Blended finance for {field:Client}', toc: true, custom_properties: { 'BwB document type': 'Proposal' } },
      blocks: [
        { type: 'heading', level: 1, text: 'Executive summary' },
        { type: 'guidance', text: 'State the decision {field:Client} needs to take, and our recommendation, in three sentences.' },
        { type: 'heading', level: 1, text: 'Context' },
        { type: 'guidance', text: 'Describe the market and the problem. Cite sources with footnotes.' },
        { type: 'heading', level: 2, text: 'Market' },
        { type: 'heading', level: 2, text: 'Problem' },
        { type: 'heading', level: 1, text: 'Approach' },
        { type: 'caption', kind: 'Figure', text: 'A caption block.' },
      ],
    }
    const { docx } = await buildDocx(content, { base, cover: coverPhoto(), icons, now: FIXED_NOW })
    save('ts-new-document.docx', docx)
  })
})

describe('inserted blocks', () => {
  it('merge into a document built on the base template', { timeout: 60_000 }, async () => {
    const blocks = everyBlock()
    const half = Math.floor(blocks.length / 2)
    const first = await buildBlockOoxml(blocks.slice(0, half), { icons })
    const second = await buildBlockOoxml(blocks.slice(half), { icons, seqStart: { Table: 2, Figure: 2 } })
    const third = await buildBlockOoxml([{ type: 'numbered', items: ['A list in a third insertion restarts at one'] }])
    for (const r of [first, second, third]) expect(readFlatOpc(r.ooxml).size).toBeGreaterThan(4)
    const docx = flatOpcToDocx(base, [first.ooxml, second.ooxml, third.ooxml])
    expect(Object.keys(unzipSync(docx))).toContain('word/media/f0_icon_leaf_navy.png')
    save('ts-inserted-blocks.docx', docx)
  })
})
