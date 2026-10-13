import { afterEach, describe, expect, it } from 'vitest'
import { TITLE_AREA } from '../../../src/brand/slideGrid.ts'
import { runDeckCheck } from '../../../src/ppt/check/engine.ts'
import { StaleDeckError, applyDeckFix, applyDeckFixes, deleteTaggedShapes, goToLocation } from '../../../src/ppt/check/fixes.ts'
import { takeDeckSnapshot } from '../../../src/ppt/check/snapshot.ts'
import type { DeckFinding, DeckFix, DeckReport } from '../../../src/ppt/check/types.ts'
import type { FakeShape } from '../../fakePpt/index.ts'
import {
  createPresentation, findShape, group, installFakePowerPoint, placeholder, run, shape, shapesTagged, slide, textBox, textFrame, textOf,
  uninstallFakePowerPoint,
} from '../../fakePpt/index.ts'

afterEach(uninstallFakePowerPoint)

const check = async (): Promise<DeckReport> => runDeckCheck(await takeDeckSnapshot())
const findingsOf = (report: DeckReport, rule: string): DeckFinding[] => report.findings.filter((f) => f.rule === rule)
const fixesOf = (report: DeckReport, rule: string): DeckFix[] => findingsOf(report, rule).flatMap((f) => (f.fix ? [f.fix] : []))
const titled = (shapes: FakeShape[], id?: string) => slide([placeholder('Title', 'An action title'), ...shapes], id ? { id } : {})

describe('wording fixes', () => {
  it('replace from the end of the text back, keeping the formatting around them', async () => {
    const box = textBox('', { id: 'b', textFrame: textFrame('', { runs: [run('The organization ', { bold: true }), run('will analyze it.')] }) })
    installFakePowerPoint(createPresentation([titled([box])]))
    const fixes = fixesOf(await check(), 'wording/us-spelling')
    expect(fixes).toHaveLength(2)
    expect(await applyDeckFixes(fixes)).toEqual({ applied: 2, skipped: 0 })
    expect(textOf(box)).toBe('The organisation will analyse it.')
    expect(box.textFrame?.runs.map((r) => [r.text, r.font.bold])).toEqual([['The organisation ', true], ['will analyse it.', false]])
  })

  it('skip text that changed after the check, rather than guess', async () => {
    const box = textBox('The organization will analyze it.', { id: 'b' })
    installFakePowerPoint(createPresentation([titled([box])]))
    const fixes = fixesOf(await check(), 'wording/us-spelling')
    box.textFrame = textFrame('Someone rewrote this organization line.')
    expect(await applyDeckFixes(fixes)).toEqual({ applied: 0, skipped: 2 })
    const first = fixes[0]
    if (!first) throw new Error('expected a fix')
    await expect(applyDeckFix(first)).rejects.toBeInstanceOf(StaleDeckError)
    expect(textOf(box)).toBe('Someone rewrote this organization line.')
  })

  it('check that each span still holds its text, and apply a repeated fix once', async () => {
    const text = 'The organization will analyze it.'
    const box = textBox(text, { id: 'b' })
    installFakePowerPoint(createPresentation([titled([box], 's1')]))
    const at = { slideId: 's1', path: ['b'], expectedText: text }
    const replace = (match: string): DeckFix => ({ ...at, kind: 'replace', start: 4, end: 16, match, replacement: 'organisation' })
    expect(await applyDeckFixes([replace('misread text')])).toEqual({ applied: 0, skipped: 1 })
    expect(textOf(box)).toBe(text)
    expect(await applyDeckFixes([replace('organization'), replace('organization')])).toEqual({ applied: 1, skipped: 1 })
    expect(textOf(box)).toBe('The organisation will analyze it.')
  })

  it('run in a few syncs however many shapes they touch', async () => {
    const boxes = Array.from({ length: 45 }, (_, i) => textBox(`Point ${i} on the organization.`, { id: `b${i}` }))
    const fake = installFakePowerPoint(createPresentation([titled(boxes)]))
    const fixes = fixesOf(await check(), 'wording/us-spelling')
    const before = fake.syncs()
    expect(await applyDeckFixes(fixes)).toEqual({ applied: 45, skipped: 0 })
    expect(fake.syncs() - before).toBeLessThanOrEqual(7)
    expect(boxes.every((b) => textOf(b).endsWith('organisation.'))).toBe(true)
  })
})

describe('brand fixes', () => {
  it('change the font, size and colour of only the text named', async () => {
    const box = textBox('', { id: 'b', textFrame: textFrame('', { runs: [run('Keep this '), run('Stray', { name: 'Arial', size: 7, color: '#FF0000' })] }) })
    installFakePowerPoint(createPresentation([titled([box])]))
    const report = await check()
    const fixes = [...fixesOf(report, 'brand/font'), ...fixesOf(report, 'brand/small-text'), ...fixesOf(report, 'brand/off-palette')]
    expect(fixes.map((f) => f.kind)).toEqual(['font', 'font-size', 'text-colour'])
    expect(await applyDeckFixes(fixes)).toEqual({ applied: 3, skipped: 0 })
    expect(box.textFrame?.runs.map((r) => [r.text, r.font.name, r.font.size, r.font.color])).toEqual([
      ['Keep this ', 'Aptos', 18, '#000000'], ['Stray', 'Aptos', 10, '#4A4A4C'],
    ])
  })

  it('recolour fills and outlines', async () => {
    const box = shape({ id: 'r', fill: { type: 'Solid', foregroundColor: '#FFC000', transparency: 0 }, line: { visible: true, color: '#5B9BD5', weight: 1, dashStyle: 'Solid', style: 'Single', transparency: 0 } })
    installFakePowerPoint(createPresentation([titled([box])]))
    const fixes = fixesOf(await check(), 'brand/off-palette')
    expect(fixes.map((f) => f.kind).sort()).toEqual(['fill-colour', 'line-colour'])
    await applyDeckFixes(fixes)
    expect(runDeckCheck(await takeDeckSnapshot()).findings.filter((f) => f.rule === 'brand/off-palette')).toEqual([])
    expect([box.fill.type, box.line.visible]).toEqual(['Solid', true])
  })

  it('reach shapes inside groups through their path', async () => {
    const inside = textBox('Grouped', { id: 'inside', font: { name: 'Calibri' } })
    installFakePowerPoint(createPresentation([titled([group([inside, shape({ id: 'other', left: 300 })], { id: 'g' })])]))
    const fixes = fixesOf(await check(), 'brand/font')
    expect(fixes).toEqual([expect.objectContaining({ path: ['g', 'inside'] })])
    expect(await applyDeckFixes(fixes)).toEqual({ applied: 1, skipped: 0 })
    expect(inside.textFrame?.runs[0]?.font.name).toBe('Aptos')
  })

  it('apply a folded fix for every shape it covers', async () => {
    const slides = Array.from({ length: 8 }, () => slide([placeholder('Title', 'A title', { font: { name: 'Calibri' } }), textBox('Body', { font: { name: 'Calibri' } })]))
    installFakePowerPoint(createPresentation(slides))
    const folded = findingsOf(await check(), 'brand/deck-font')[0]
    expect(folded?.fix?.kind).toBe('batch')
    expect(await applyDeckFixes(folded?.fix ? [folded.fix] : [])).toEqual({ applied: 16, skipped: 0 })
    const names = slides.flatMap((s) => s.shapes.map((x) => x.textFrame?.runs[0]?.font.name))
    expect(names).toEqual(Array.from({ length: 8 }, () => ['Candara', 'Aptos']).flat())
  })
})

describe('layout fixes', () => {
  it('line up a jumping title, unless it moved since the check', async () => {
    const jumping = placeholder('Title', 'Jumps', { id: 'j', top: 40, height: 80 })
    const deck = createPresentation([titled([]), titled([]), titled([]), slide([jumping])])
    installFakePowerPoint(deck)
    const fix = fixesOf(await check(), 'layout/jumping-title')[0]
    if (!fix) throw new Error('expected a fix')
    jumping.left = 50
    expect(await applyDeckFixes([fix])).toEqual({ applied: 0, skipped: 1 })
    jumping.left = 40
    await applyDeckFix(fix)
    expect([jumping.left, jumping.top, jumping.width, jumping.height]).toEqual([TITLE_AREA.left, TITLE_AREA.top, TITLE_AREA.width, TITLE_AREA.height])
  })

  it('move a shape back onto the slide, and delete one parked off it', async () => {
    const edge = textBox('Edge', { id: 'edge', left: 900, top: 200, width: 100 })
    const parked = textBox('Parked', { id: 'parked', left: 1000, top: 200 })
    const deck = createPresentation([titled([edge, parked])])
    installFakePowerPoint(deck)
    const report = await check()
    await applyDeckFixes([...fixesOf(report, 'layout/off-edge'), ...fixesOf(report, 'layout/off-slide')])
    expect(edge.left).toBe(860)
    expect(findShape(deck, 'parked')).toBeUndefined()
  })

  it('delete empty placeholders and text boxes, but not one filled since the check', async () => {
    const emptyBody = placeholder('Content', '', { id: 'body' })
    const emptyPicture = placeholder('Picture', '', { id: 'pic' })
    const emptyBox = textBox('', { id: 'box', top: 300 })
    const filledLater = textBox('', { id: 'later', top: 400 })
    const deck = createPresentation([titled([emptyBody, emptyPicture, emptyBox, filledLater])])
    installFakePowerPoint(deck)
    const report = await check()
    const fixes = [...fixesOf(report, 'layout/empty-placeholder'), ...fixesOf(report, 'layout/empty-text-box')]
    expect(fixes).toHaveLength(4)
    filledLater.textFrame = textFrame('Typed since')
    emptyPicture.textFrame = null
    expect(await applyDeckFixes(fixes)).toEqual({ applied: 2, skipped: 2 })
    expect(['body', 'pic', 'box', 'later'].map((id) => Boolean(findShape(deck, id)))).toEqual([false, true, false, true])
  })

  it('skip shapes that are gone', async () => {
    const deck = createPresentation([titled([textBox('The organization', { id: 'gone' })], 's1')])
    installFakePowerPoint(deck)
    const fixes = fixesOf(await check(), 'wording/us-spelling')
    deck.slides[0]?.shapes.pop()
    expect(await applyDeckFixes(fixes)).toEqual({ applied: 0, skipped: 1 })
  })
})

describe('before-you-send fixes', () => {
  const note = (text: string): FakeShape => textBox(text, { tags: { BWB_NOTE: '2026-10-13', BWB_ELEMENT: 'note' } })
  const sticker = (label: string): FakeShape => textBox(label.toUpperCase(), { tags: { BWB_STICKER: label, BWB_ELEMENT: 'sticker' }, top: 5, left: 800 })

  it('remove every review note, and only the stickers with the label named', async () => {
    const deck = createPresentation([
      titled([note('Note: check this'), sticker('Draft')]),
      titled([note('Note: and this'), sticker('Draft'), sticker('Confidential')]),
    ])
    installFakePowerPoint(deck)
    const report = await check()
    await applyDeckFixes(fixesOf(report, 'send/review-notes'))
    expect(shapesTagged(deck, 'BWB_NOTE')).toEqual([])
    const draft = findingsOf(report, 'send/sticker').find((f) => f.title.startsWith('“DRAFT”'))
    await applyDeckFixes(draft?.fix ? [draft.fix] : [])
    expect(shapesTagged(deck, 'BWB_STICKER').map((s) => s.tags.get('BWB_STICKER'))).toEqual(['Confidential'])
  })

  it('delete tagged shapes inside tagged groups without tripping over them', async () => {
    const deck = createPresentation([titled([group([note('Inner note'), textBox('x')], { id: 'g', tags: { BWB_NOTE: 'grouped' } })])])
    installFakePowerPoint(deck)
    expect(await deleteTaggedShapes('BWB_NOTE')).toBe(2)
    expect(findShape(deck, 'g')).toBeUndefined()
  })

  it('update the trackers as one fix', async () => {
    const divider = (title: string) => slide([placeholder('Title', title)], { layoutId: 'layout-section' })
    const tracker = textBox('Context | Old name', { id: 'tr', tags: { BWB_TRACKER: '1:Context' }, top: 8 })
    const deck = createPresentation([divider('Context'), titled([tracker]), divider('Options')])
    installFakePowerPoint(deck)
    const fixes = fixesOf(await check(), 'send/trackers')
    expect(fixes).toEqual([{ kind: 'update-trackers' }])
    expect(await applyDeckFixes([...fixes, ...fixes])).toEqual({ applied: 1, skipped: 0 })
    expect(textOf(tracker)).toBe('Context | Options')
  })
})

describe('going to a finding', () => {
  function deck() {
    const box = textBox('The organization', { id: 'b' })
    const inner = textBox('In a group', { id: 'inner' })
    return createPresentation([slide([placeholder('Title', 'One')], { id: 's1' }), slide([box, group([inner, shape()], { id: 'g' })], { id: 's2' })])
  }

  it('selects the slide, the shape or the text', async () => {
    const model = deck()
    installFakePowerPoint(model)
    await goToLocation({ kind: 'slide', slideId: 's2', slide: 2 })
    expect(model.selection).toEqual({ slideIds: ['s2'], shapeIds: [], textRange: null })
    await goToLocation({ kind: 'shape', slideId: 's2', slide: 2, path: ['b'] })
    expect(model.selection.shapeIds).toEqual(['b'])
    await goToLocation({ kind: 'shape', slideId: 's2', slide: 2, path: ['b'], start: 4, end: 16 })
    expect(model.selection.textRange).toEqual({ shapeId: 'b', start: 4, length: 12 })
    await goToLocation({ kind: 'shape', slideId: 's2', slide: 2, path: ['g', 'inner'] })
    expect(model.selection.shapeIds).toEqual(['g'])
  })

  it('says when the slide has gone, or PowerPoint is too old', async () => {
    installFakePowerPoint(deck())
    await expect(goToLocation({ kind: 'slide', slideId: 'gone', slide: 9 })).rejects.toBeInstanceOf(StaleDeckError)
    uninstallFakePowerPoint()
    installFakePowerPoint(deck(), { apiVersion: '1.4' })
    await expect(goToLocation({ kind: 'slide', slideId: 's1', slide: 1 })).rejects.toThrow('Going to a slide needs a newer PowerPoint')
  })
})
