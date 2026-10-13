import { DOMParser } from '@xmldom/xmldom'
import { afterEach, beforeAll, describe, expect, it } from 'vitest'
import { TEMPLATE_MARKER_STYLES } from '../../src/brand/styles.ts'
import { qaNote, runChecks } from '../../src/check/engine.ts'
import { takeSnapshot } from '../../src/word/snapshot.ts'
import { createModel, font, para, section } from '../fakeWord/model.ts'
import type { FakeModel } from '../fakeWord/model.ts'
import { installFakeWord, uninstallFakeWord } from '../fakeWord/install.ts'

beforeAll(() => {
  // The deep scan parses the body's WordprocessingML with the browser's DOMParser.
  Object.assign(globalThis, { DOMParser })
})
afterEach(uninstallFakeWord)

function messyModel(): FakeModel {
  const model = createModel([
    para('Blended finance for adaptation', 'Title', { styleBuiltIn: 'Title' }),
    para('1. Introduction', 'heading 1'),
    para('The organization will analyze the data from 2022\u20132025.'),
    para('Results', 'heading 3'),
    para(''),
    para(''),
    para('Figures to confirm with the client', 'Normal', { font: font({ highlightColor: 'Yellow' }) }),
    para('A paragraph someone pasted in Arial.', 'Normal', { font: font({ name: 'Arial' }) }),
  ])
  model.trackedChanges = 2
  model.comments = 1
  model.tables.push({ style: 'Table Grid', rows: [['Country', 'Size (USD m)'], ['Angola', '250']] })
  model.pictures.push({ altTextTitle: '', altTextDescription: '' })
  model.sections = [section('', 'Bankers without Boundaries', 1)]
  model.url = 'https://bwb.sharepoint.com/Shared%20Documents/Report%20-%20Copy.docx'
  model.properties.author = 'J. Smith'
  return model
}

function cleanModel(): FakeModel {
  const model = createModel([
    para('Blended finance for adaptation', 'Title', { styleBuiltIn: 'Title' }),
    para('Introduction', 'heading 1'),
    para('The organisation will analyse the data from 2022 to 2025.'),
    para('Scope', 'heading 2'),
    para('Bankers without Boundaries (BwB) reviewed 12 programmes.'),
    para('Disclaimer', 'BwB Disclaimer Heading'),
    para('BwB (FRN: 953637) is authorised and regulated by the UK Financial Conduct Authority.'),
  ])
  TEMPLATE_MARKER_STYLES.forEach((s) => model.styles.add(s))
  model.properties = { title: 'Blended finance for adaptation', subject: '', author: 'Bankers without Boundaries', company: 'Bankers without Boundaries' }
  model.url = 'https://bwb.sharepoint.com/Shared%20Documents/20261013_NZC_D3.2_Blended%20finance_v04.docx'
  return model
}

describe('the Pre-QA check on a messy document', () => {
  it('finds the problems a QA reviewer would raise', async () => {
    installFakeWord(messyModel())
    const report = runChecks(await takeSnapshot({ deepScan: true }))
    const rules = new Set(report.findings.map((f) => f.rule))
    for (const rule of [
      'send/tracked-changes', 'send/comments', 'send/no-title', 'send/properties', 'send/page-numbers',
      'structure/typed-number', 'structure/skipped-level', 'structure/empty-paragraph', 'structure/no-disclaimer',
      'writing/us-spelling', 'brand/highlight', 'brand/font', 'brand/not-template',
      'tables/style', 'tables/no-caption', 'tables/alt-text',
    ]) {
      expect(rules, rule).toContain(rule)
    }
    expect([...rules].some((r) => r.startsWith('send/file-name-'))).toBe(true)
    expect(report.ready).toBe(false)
    expect(report.counts.major).toBeGreaterThanOrEqual(2)
    expect(report.stats).toMatchObject({ tables: 1, pictures: 1, headings: 2 })
  })

  it('sorts findings from blocker to copyedit, with suggestions last within each severity', async () => {
    installFakeWord(messyModel())
    const { findings } = runChecks(await takeSnapshot({ deepScan: true }))
    const order = { blocker: 0, major: 1, minor: 2, copyedit: 3 }
    for (let i = 1; i < findings.length; i++) {
      const a = findings[i - 1]
      const b = findings[i]
      if (!a || !b) continue
      expect(order[a.severity]).toBeLessThanOrEqual(order[b.severity])
      if (a.severity === b.severity) expect(Number(Boolean(a.advisory))).toBeLessThanOrEqual(Number(Boolean(b.advisory)))
    }
  })

  it('writes a QA note that lists the major items', async () => {
    installFakeWord(messyModel())
    const report = runChecks(await takeSnapshot({ deepScan: true }))
    const note = qaNote(report, 'Blended finance for adaptation', new Date(2026, 9, 13))
    expect(note).toContain('BwB Pre-QA check: Blended finance for adaptation, 13 October 2026')
    expect(note).toContain('Not ready yet. Major items:')
    expect(note).toMatch(/- 2 tracked changes still in the document/)
  })
})

describe('the Pre-QA check on a clean BwB document', () => {
  it('has no blocker or major items, and so is ready for QA', async () => {
    installFakeWord(cleanModel())
    const report = runChecks(await takeSnapshot({ deepScan: true }))
    const serious = report.findings.filter((f) => !f.advisory && (f.severity === 'blocker' || f.severity === 'major'))
    expect(serious).toEqual([])
    expect(report.ready).toBe(true)
    expect(report.findings.filter((f) => f.group === 'writing' && !f.advisory)).toEqual([])
    expect(report.acronyms.find((a) => a.acronym === 'BwB')).toBeDefined()
  })
})
