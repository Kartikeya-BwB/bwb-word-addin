import { afterEach, beforeEach, describe, expect, it, vi } from 'vitest'
import ribbon from '../../src/commands/ribbon-powerpoint.json'
import { TAGS } from '../../src/brand/slideGrid.ts'
import { PPT_HANDLERS } from '../../src/ppt/handlers.ts'
import { selectionTracker } from '../../src/ppt/selection.ts'
import type { FakePresentation } from '../fakePpt/index.ts'
import { createPresentation, group, installFakePowerPoint, select, shape, slide, textBox, uninstallFakePowerPoint } from '../fakePpt/index.ts'
import { iconFetch } from './helpers.ts'

/** The ribbon buttons the slide tools own (agenda, trackers, sections). */
const SLIDE_TOOLS = ['pptAgendaInsert', 'pptTrackersAdd', 'pptTrackersUpdate', 'pptTrackersRemove', 'pptSectionMark']
/** Buttons whose handlers live elsewhere: the slide tools, and New BwB deck (src/ppt/newDeck.ts). */
const ELSEWHERE = [...SLIDE_TOOLS, 'pptNewDeck']

interface Entry { fn?: string; items?: Entry[] }
const ribbonFns = (): string[] =>
  (ribbon as { groups: Array<{ controls: Entry[] }> }).groups
    .flatMap((g) => g.controls)
    .flatMap((c) => [c, ...(c.items ?? [])])
    .flatMap((e) => (e.fn ? [e.fn] : []))

beforeEach(() => {
  vi.stubGlobal('fetch', iconFetch)
})

afterEach(async () => {
  await selectionTracker.stop()
  selectionTracker.clearAnchor()
  selectionTracker.apply({ slideId: null, shapes: [] })
  uninstallFakePowerPoint()
  vi.unstubAllGlobals()
})

describe('the PowerPoint ribbon handlers', () => {
  it('cover every function button except the slide tools and New BwB deck, and nothing else', () => {
    const fns = ribbonFns().filter((fn) => !ELSEWHERE.includes(fn))
    expect(Object.keys(PPT_HANDLERS).sort()).toEqual([...fns].sort())
    for (const fn of SLIDE_TOOLS) expect(ribbonFns()).toContain(fn)
  })
})

/**
 * A slide with three selected shapes between two others, laid out roughly, with US spelling, an Arial text box that
 * grows to fit, and a sticker and a review note on another slide: every handler has something to change.
 */
function fixture(selection: 'three' | 'two' | 'group'): FakePresentation {
  const shapes = [
    shape({ id: 'back' }),
    shape({ id: 'a', left: 60, top: 120, width: 120, height: 60, text: 'The organization' }),
    shape({ id: 'b', left: 330, top: 135, width: 90, height: 50, text: 'Plan' }),
    textBox('Analyze', { id: 't', left: 170, top: 270, width: 150, height: 40, font: { name: 'Arial' } }),
    shape({ id: 'front' }),
    group([shape({ id: 'm1', left: 500, top: 300 }), shape({ id: 'm2', left: 620, top: 320 })], { id: 'g' }),
  ]
  const other = slide([
    shape({ id: 'old-sticker', tags: { [TAGS.sticker]: 'Draft', [TAGS.element]: 'sticker' } }),
    shape({ id: 'old-note', tags: { [TAGS.note]: '2026-10-01', [TAGS.element]: 'note' } }),
  ], { id: 's2' })
  const d = createPresentation([slide(shapes, { id: 's1' }), other])
  select(d, selection === 'three' ? ['a', 'b', 't'] : selection === 'two' ? ['a', 'b'] : ['g'], 's1')
  return d
}

const SELECTION: Record<string, 'two' | 'group'> = { pptSwap: 'two', pptUngroup: 'group' }

const snapshot = (d: FakePresentation): string => JSON.stringify(d.slides, (_k, v: unknown) => (v instanceof Map ? [...v] : v))

describe.each(Object.keys(PPT_HANDLERS).sort())('%s', (fn) => {
  it('runs against PowerPoint and changes the deck', async () => {
    const d = fixture(SELECTION[fn] ?? 'three')
    installFakePowerPoint(d)
    const before = snapshot(d)
    const handler = PPT_HANDLERS[fn]
    expect(handler).toBeTypeOf('function')
    const result = await handler?.()
    expect(typeof result).toBe('string')
    expect(String(result)).not.toMatch(/[\u2014!]/)
    expect(snapshot(d), String(result)).not.toBe(before)
  })
})
