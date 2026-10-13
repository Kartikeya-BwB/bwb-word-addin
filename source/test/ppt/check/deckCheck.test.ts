import { afterEach, describe, expect, it } from 'vitest'
import { TITLE_AREA } from '../../../src/brand/slideGrid.ts'
import { deckBulkFixes, runDeckCheck } from '../../../src/ppt/check/engine.ts'
import { applyDeckFix, applyDeckFixes } from '../../../src/ppt/check/fixes.ts'
import { takeDeckSnapshot } from '../../../src/ppt/check/snapshot.ts'
import type { DeckReport } from '../../../src/ppt/check/types.ts'
import type { FakePresentation } from '../../fakePpt/index.ts'
import {
  createPresentation, findShape, installFakePowerPoint, picture, placeholder, run, shape, shapesTagged, slide, textBox, textFrame, textOf,
  uninstallFakePowerPoint,
} from '../../fakePpt/index.ts'

afterEach(uninstallFakePowerPoint)

const check = async (): Promise<DeckReport> => runDeckCheck(await takeDeckSnapshot())
const rules = (report: DeckReport): Set<string> => new Set(report.findings.map((f) => f.rule))

/** A deck with the problems a QA reviewer finds in a first draft. */
function messyDeck(): FakePresentation {
  const titled = (text: string, id: string, over = {}) => placeholder('Title', text, { id, ...over })
  return createPresentation([
    slide([placeholder('CenterTitle', 'Blended finance for adaptation'), placeholder('Subtitle', 'Board meeting')], { layoutId: 'layout-title' }),
    slide([
      titled('The organization will analyze adaptation finance', 't2'),
      placeholder('Content', '', { id: 'body2', textFrame: textFrame('', { runs: [run('Flows reached USD 63bn. '), run('Small print', { name: 'Arial', size: 9 })] }) }),
      textBox('Note: check the 2022 figure', { id: 'note', tags: { BWB_NOTE: '2026-10-13', BWB_ELEMENT: 'note' }, top: 30, left: 600 }),
      picture({ id: 'photo', left: 600, top: 200 }),
      textBox('', { id: 'empty', top: 300 }),
    ]),
    slide([
      titled('Costs rise with every year of delay', 't3', { top: 44 }),
      placeholder('Content', '[Insert chart]', { id: 'body3' }),
      textBox('DRAFT', { id: 'sticker', tags: { BWB_STICKER: 'Draft', BWB_ELEMENT: 'sticker' }, top: 5, left: 860, width: 60 }),
    ]),
    slide([
      titled('Guarantees crowd in private lenders', 't4'),
      placeholder('Content', 'Three pilots closed in 2025.'),
      shape({ id: 'orange', left: 600, top: 200, fill: { type: 'Solid', foregroundColor: '#ED7D31', transparency: 0 } }),
    ]),
    slide([titled('Pooling projects cuts transaction costs', 't5'), placeholder('Content', 'Pools lower the cost per deal.')]),
    slide([picture({ id: 'only-photo', altTextDescription: 'Mangroves at low tide' })]),
  ])
}

describe('the deck check on a messy deck', () => {
  it('finds the problems, fixes the safe ones in one go, and leaves the rest for the author', async () => {
    const deck = messyDeck()
    installFakePowerPoint(deck)
    const first = await check()
    for (const rule of [
      'send/review-notes', 'send/sticker', 'wording/placeholder', 'wording/us-spelling', 'brand/font', 'brand/small-text',
      'brand/off-palette', 'layout/jumping-title', 'layout/empty-text-box', 'slides/no-title', 'slides/alt-text',
    ]) expect(rules(first), rule).toContain(rule)
    expect(first.ready).toBe(false)
    expect(first.stats).toMatchObject({ slides: 6, pictures: 2 })

    const safe = deckBulkFixes(first.findings)
    expect(safe.map((f) => f.rule).sort()).toEqual(['brand/font', 'layout/empty-text-box', 'layout/jumping-title', 'wording/us-spelling', 'wording/us-spelling'])
    expect(await applyDeckFixes(safe.flatMap((f) => (f.fix ? [f.fix] : [])))).toEqual({ applied: 5, skipped: 0 })

    expect(textOf(findShape(deck, 't2') ?? shape())).toBe('The organisation will analyse adaptation finance')
    expect(findShape(deck, 'body2')?.textFrame?.runs.map((r) => [r.text, r.font.name, r.font.size])).toEqual([
      ['Flows reached USD 63bn. ', 'Aptos', 18], ['Small print', 'Aptos', 9],
    ])
    expect(findShape(deck, 't3')?.top).toBe(TITLE_AREA.top)
    expect(findShape(deck, 'empty')).toBeUndefined()

    const second = await check()
    expect(deckBulkFixes(second.findings)).toEqual([])
    expect([...rules(second)].sort()).toEqual([
      'brand/off-palette', 'brand/small-text', 'send/review-notes', 'send/sticker', 'slides/alt-text', 'slides/no-title', 'wording/placeholder',
    ])

    // The rest, one at a time, as an author would.
    for (const rule of ['send/review-notes', 'brand/small-text', 'brand/off-palette', 'send/sticker']) {
      for (const f of second.findings.filter((x) => x.rule === rule)) if (f.fix) await applyDeckFix(f.fix)
    }
    expect(findShape(deck, 'orange')?.fill.foregroundColor).not.toBe('#ED7D31')
    expect(shapesTagged(deck, 'BWB_NOTE')).toEqual([])
    expect(shapesTagged(deck, 'BWB_STICKER')).toEqual([])
    const third = await check()
    expect([...rules(third)].sort()).toEqual(['slides/alt-text', 'slides/no-title', 'wording/placeholder'])
    expect(third.counts).toEqual({ blocker: 0, major: 2, minor: 1, copyedit: 0 })
  })
})
