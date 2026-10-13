import { describe, expect, it } from 'vitest'
import { deckBulkFixes, deckLimits, deckQaNote, isBulkSafe, runDeckCheck } from '../../../src/ppt/check/engine.ts'
import { flattenFixes, orderReplacements, sameBox } from '../../../src/ppt/check/fixes.ts'
import type { DeckFinding, DeckFix } from '../../../src/ppt/check/types.ts'
import { body, can, deck, font, picture, slide, source, text, title } from './build.ts'

function messyDeck() {
  return deck([
    slide([title('Blended finance for adaptation', { placeholder: 'CenterTitle', top: 200 })], { layout: 'Title Slide' }),
    slide([title('Context')], { tags: { BWB_LAYOUT: 'divider' } }),
    slide([
      title('The organization will analyze adaptation finance'),
      body('Figures in Arial at 9 pt', { font: font({ name: 'Arial', size: 9 }) }),
      text('Check the 2022-2025 figures', { tags: { BWB_NOTE: '1' }, top: 10 }),
      picture(),
    ]),
    slide([title('Second point', { top: 40 }), body('[Insert chart]'), text('', {})]),
    slide([title('Third point'), body('Words.'), source('Source: BwB')]),
    slide([picture()]),
  ])
}

function cleanDeck() {
  return deck([
    slide([title('Blended finance for adaptation', { placeholder: 'CenterTitle', top: 200 })], { layout: 'Title Slide' }),
    slide([title('Context')], { tags: { BWB_LAYOUT: 'divider' } }),
    slide([title('Adaptation finance falls short of need'), body('Flows reached USD 63bn in 2022.'), source('Source: BwB analysis')]),
    slide([title('Private capital can close part of the gap'), body('Guarantees crowd in lenders.')]),
    slide([title('Three steps would help'), body('Pool projects, share risk and report results.')]),
  ])
}

describe('the Pre-QA check for decks', () => {
  it('finds what a QA reviewer would raise on a messy deck', () => {
    const report = runDeckCheck(messyDeck())
    const found = new Set(report.findings.map((f) => f.rule))
    for (const rule of [
      'send/review-notes', 'wording/placeholder', 'brand/font', 'brand/small-text', 'layout/jumping-title',
      'layout/empty-text-box', 'wording/us-spelling', 'slides/no-title', 'slides/alt-text',
    ]) expect(found, rule).toContain(rule)
    expect(report.ready).toBe(false)
    expect(report.counts.major).toBe(3)
  })

  it('passes a clean deck', () => {
    const report = runDeckCheck(cleanDeck())
    expect(report.findings.map((f) => `${f.rule}: ${f.title}`)).toEqual([])
    expect(report.ready).toBe(true)
    expect(report.stats).toEqual({ slides: 5, shapes: 9, words: 43, pictures: 0 })
  })

  it('sorts by severity, then suggestions last, then slide', () => {
    const report = runDeckCheck(messyDeck())
    const order = { blocker: 0, major: 1, minor: 2, copyedit: 3 }
    const keys = report.findings.map((f) => [order[f.severity], Number(Boolean(f.advisory)), f.slide ?? 0])
    const sorted = [...keys].sort((a, b) => (a[0] ?? 0) - (b[0] ?? 0) || (a[1] ?? 0) - (b[1] ?? 0) || (a[2] ?? 0) - (b[2] ?? 0))
    expect(keys).toEqual(sorted)
    const counted = report.findings.filter((f) => !f.advisory).length
    expect(report.counts.blocker + report.counts.major + report.counts.minor + report.counts.copyedit).toBe(counted)
  })

  it('says which checks this PowerPoint could not run', () => {
    const old = deck([], { can: can({ placeholders: false, groups: false, altText: false, pageSize: false, selection: false, hidden: false, background: false, links: false, markers: false }), partial: true })
    const limits = deckLimits(old)
    expect(limits).toHaveLength(5)
    expect(limits.join(' ')).toMatch(/API 1\.8.*API 1\.10.*API 1\.10.*API 1\.5/)
    expect(deckLimits(deck([]))).toEqual([])
    expect(runDeckCheck(old).limits).toEqual(limits)
  })
})

describe('fixing in bulk', () => {
  const f = (fix: DeckFix, over: Partial<DeckFinding> = {}): DeckFinding => ({ id: fix.kind, rule: 'r', group: 'brand', severity: 'minor', title: 't', fix, ...over })
  const ref = { slideId: '256#1', path: ['2'] }
  const box = { left: 0, top: 0, width: 1, height: 1 }

  it('runs only fixes that change nothing but what they name', () => {
    const safe: DeckFix[] = [
      { ...ref, kind: 'replace', expectedText: 'a', start: 0, end: 1, match: 'a', replacement: 'b' },
      { ...ref, kind: 'font', expectedText: 'a', spans: [], font: 'Aptos' },
      { ...ref, kind: 'align-title', from: box, to: box },
      { ...ref, kind: 'delete-empty', textual: true },
      { kind: 'update-trackers' },
    ]
    const unsafe: DeckFix[] = [
      { ...ref, kind: 'text-colour', expectedText: 'a', spans: [], colour: '022338' },
      { ...ref, kind: 'fill-colour', colour: '022338' },
      { ...ref, kind: 'font-size', expectedText: 'a', spans: [], size: 10 },
      { ...ref, kind: 'move-onto-slide', from: box, to: box },
      { ...ref, kind: 'delete-shape', from: box },
      { kind: 'delete-tagged', tag: 'BWB_NOTE' },
      { kind: 'batch', fixes: [] },
    ]
    for (const fix of safe) expect(isBulkSafe(fix), fix.kind).toBe(true)
    for (const fix of unsafe) expect(isBulkSafe(fix), fix.kind).toBe(false)
    expect(isBulkSafe({ kind: 'batch', fixes: safe })).toBe(true)
    expect(isBulkSafe({ kind: 'batch', fixes: [...safe, ...unsafe.slice(0, 1)] })).toBe(false)
    const findings = [...safe.map((x) => f(x)), ...unsafe.map((x) => f(x)), f(safe[0] as DeckFix, { advisory: true })]
    expect(deckBulkFixes(findings)).toHaveLength(safe.length)
  })

  it('flattens batches and orders replacements from the end of the text back, skipping overlaps', () => {
    const r = (start: number, end: number): Extract<DeckFix, { kind: 'replace' }> =>
      ({ ...ref, kind: 'replace', expectedText: 'x', start, end, match: 'x', replacement: 'y' })
    expect(flattenFixes([{ kind: 'batch', fixes: [r(0, 1), { kind: 'batch', fixes: [r(2, 3)] }] }, r(4, 5)]).map((x) => x.kind === 'replace' && x.start)).toEqual([0, 2, 4])
    const { apply, skipped } = orderReplacements([r(0, 4), r(10, 12), r(3, 6), r(10, 12), r(20, 20)])
    expect(apply.map((x) => x.start)).toEqual([10, 3])
    expect(skipped.map((x) => x.start)).toEqual([20, 10, 0])
  })

  it('compares positions within half a point', () => {
    expect(sameBox({ left: 40, top: 28, width: 880, height: 64 }, { left: 40.4, top: 27.6, width: 880, height: 64 })).toBe(true)
    expect(sameBox({ left: 40, top: 28, width: 880, height: 64 }, { left: 41, top: 28, width: 880, height: 64 })).toBe(false)
  })
})

describe('the QA note', () => {
  it('summarises the deck and lists the major items with their slides', () => {
    const report = runDeckCheck(messyDeck())
    const note = deckQaNote(report, 'Adaptation finance', new Date(2026, 9, 13))
    const lines = note.split('\n')
    expect(lines[0]).toBe('BwB Pre-QA check: Adaptation finance, 13 October 2026')
    expect(lines[1]).toMatch(/^6 slides, \d+ words, 2 pictures$/)
    expect(lines[3]).toBe('Not ready yet. Major items:')
    expect(lines).toContain('- Slide 3: 1 review note left on 1 slide')
    expect(deckQaNote(runDeckCheck(cleanDeck()), '', new Date(2026, 9, 13)).split('\n')[3]).toBe('No blocker or major items: ready for peer review and QA.')
  })
})
