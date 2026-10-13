import { readFileSync } from 'node:fs'
import { join } from 'node:path'
import { afterEach, describe, expect, it } from 'vitest'
import {
  HEADING_STYLES, LIST_STYLES, TABLE_STYLES, TEMPLATE_MARKER_STYLES, TEXT_STYLES, UNNUMBERED_HEADING_STYLES,
} from '../../src/brand/styles.ts'
import { addBwbStyles, applyListStyle, applyParagraphStyle, hasBwbStyles, restyleBuiltInStyles } from '../../src/word/styles.ts'
import { listTemplateStyles, styleCarrierBody } from '../../src/word/flatOpc.ts'
import { createModel, para } from '../fakeWord/model.ts'
import { installFakeWord, uninstallFakeWord } from '../fakeWord/install.ts'

afterEach(uninstallFakeWord)

const lead = TEXT_STYLES.find((s) => s.id === 'BwBLead')
const heading1 = HEADING_STYLES[0]

/** Style names to ids in the template's styles.xml. */
function templateIds(): Map<string, string> {
  const xml = readFileSync(join(import.meta.dirname, '../../src/docx/template/styles.xml'), 'utf8')
  const out = new Map<string, string>()
  for (const m of xml.matchAll(/<w:style\b[^>]*w:styleId="([^"]+)"[^>]*>\s*<w:name w:val="([^"]+)"/g)) {
    if (m[1] && m[2]) out.set(m[2], m[1])
  }
  return out
}

describe('the template styles the carrier brings in', () => {
  it('cover every custom style the add-in applies by name, and the marker styles', () => {
    const ids = templateIds()
    const carrier = styleCarrierBody()
    const names = [
      ...TEXT_STYLES.filter((t) => !t.builtIn), ...UNNUMBERED_HEADING_STYLES, ...LIST_STYLES,
    ].map((s) => s.name).concat(TEMPLATE_MARKER_STYLES, Object.values(TABLE_STYLES).map((t) => t.name))
    for (const name of names) {
      const id = ids.get(name)
      expect(id, name).toBeDefined()
      expect(carrier).toContain(`w:val="${id}"`)
    }
    expect(listTemplateStyles().length).toBeGreaterThan(40)
  })
})

describe('addBwbStyles', () => {
  it('adds the styles and leaves the text exactly as it was', async () => {
    const model = createModel([para('First'), para('Second')])
    installFakeWord(model)
    expect(await hasBwbStyles()).toBe(false)
    await addBwbStyles()
    expect(model.paragraphs.map((p) => p.text)).toEqual(['First', 'Second'])
    expect(model.tables).toHaveLength(0)
    expect(await hasBwbStyles()).toBe(true)
  })

  it('removes the empty paragraph Word can leave behind', async () => {
    const model = createModel([para('Only')], { ooxmlQuirk: true })
    installFakeWord(model)
    await addBwbStyles()
    expect(model.paragraphs.map((p) => p.text)).toEqual(['Only'])
  })
})

describe('applyParagraphStyle', () => {
  it('uses the built-in identifier for headings, so it works in any Word language', async () => {
    const p = para('Introduction')
    const model = createModel([p])
    installFakeWord(model)
    const result = await applyParagraphStyle(heading1 ?? { name: 'heading 1', builtIn: 'Heading1' })
    expect(p.styleBuiltIn).toBe('Heading1')
    expect(result).toEqual({ paragraphs: 1, importedStyles: false })
  })

  it('brings in the BwB styles when the document lacks them, then applies the style', async () => {
    const p = para('An opening sentence.')
    const model = createModel([p])
    installFakeWord(model)
    const result = await applyParagraphStyle(lead ?? { name: 'BwB Lead' })
    expect(p.style).toBe('BwB Lead')
    expect(result.importedStyles).toBe(true)
  })

  it('recovers on older Word that cannot list styles, by importing after the first failure', async () => {
    const p = para('An opening sentence.')
    const model = createModel([p], { apiVersion: 1.4 })
    installFakeWord(model)
    const result = await applyParagraphStyle(lead ?? { name: 'BwB Lead' })
    expect(p.style).toBe('BwB Lead')
    expect(result.importedStyles).toBe(true)
  })

  it('applies a list style to every selected paragraph', async () => {
    const items = [para('One'), para('Two'), para('Three')]
    const model = createModel(items, { selection: items })
    model.styles.add('BwB Bullet Green Square')
    installFakeWord(model)
    const square = LIST_STYLES[0]
    if (!square) throw new Error('no list styles')
    const result = await applyListStyle(square)
    expect(items.map((p) => p.style)).toEqual(Array(3).fill('BwB Bullet Green Square'))
    expect(result).toEqual({ paragraphs: 3, importedStyles: false })
  })
})

describe('restyleBuiltInStyles', () => {
  it('sets Normal, Heading 1 to 4 and Caption to the BwB specification', async () => {
    const model = createModel()
    installFakeWord(model)
    const changed = await restyleBuiltInStyles()
    expect(changed).toEqual(['Normal', 'Heading 1', 'Heading 2', 'Heading 3', 'Heading 4', 'Caption'])
    const h1 = model.styleUpdates.find((u) => u.name === 'Heading 1' && u.font)
    expect(h1?.font).toMatchObject({ name: 'Candara', size: 18, color: '#022338', bold: true })
    const normal = model.styleUpdates.find((u) => u.name === 'Normal' && u.paragraph)
    expect(normal?.paragraph).toMatchObject({ alignment: 'Justified', spaceAfter: 8, lineSpacing: 16 })
  })

  it('skips styles the document does not have', async () => {
    const model = createModel()
    model.styles.delete('Caption')
    installFakeWord(model)
    expect(await restyleBuiltInStyles()).not.toContain('Caption')
  })
})
