import { afterEach, describe, expect, it, vi } from 'vitest'
import { SelectionTracker, anchorHint, emptyState, loadState, nextOrder, selectionTracker } from '../../src/ppt/selection.ts'
import type { SelectionRead, SelectionState } from '../../src/ppt/selection.ts'
import { clickShapes, createPresentation, installFakePowerPoint, shape, slide, uninstallFakePowerPoint } from '../fakePpt/index.ts'
import { memoryStorage } from './helpers.ts'

afterEach(async () => {
  await selectionTracker.stop()
  selectionTracker.clearAnchor()
  selectionTracker.apply({ slideId: null, shapes: [] })
  uninstallFakePowerPoint()
  vi.unstubAllGlobals()
})

const read = (slideId: string | null, ...ids: string[]): SelectionRead => ({ slideId, shapes: ids.map((id) => ({ id, name: `Shape ${id}` })) })

describe('the order of selection', () => {
  it('keeps the shapes still selected in the order they were picked and adds new ones after', () => {
    expect(nextOrder([], ['a', 'b'])).toEqual(['a', 'b'])
    expect(nextOrder(['c', 'a'], ['a', 'b', 'c'])).toEqual(['c', 'a', 'b'])
    expect(nextOrder(['c', 'a', 'b'], ['a', 'b'])).toEqual(['a', 'b'])
    expect(nextOrder(['a'], [])).toEqual([])
  })

  it('is built from successive reads, and starts afresh on another slide', () => {
    const tracker = new SelectionTracker(async () => read(null))
    tracker.apply(read('s1', 'c'))
    tracker.apply(read('s1', 'a', 'c'))
    tracker.apply(read('s1', 'a', 'b', 'c'))
    expect(tracker.current.order).toEqual(['c', 'a', 'b'])
    expect(tracker.current.names.b).toBe('Shape b')
    tracker.apply(read('s2', 'a', 'b'))
    expect(tracker.current.order).toEqual(['a', 'b'])
  })

  it('tells listeners about every change until they unsubscribe', () => {
    const tracker = new SelectionTracker(async () => read(null))
    const seen: string[][] = []
    const off = tracker.subscribe((s) => seen.push(s.order))
    tracker.apply(read('s1', 'a'))
    off()
    tracker.apply(read('s1', 'a', 'b'))
    expect(seen).toEqual([['a']])
  })

  it('reads again when the selection changes during a read, so the last change wins', async () => {
    const reads = [read('s1', 'a'), read('s1', 'a', 'b')]
    let calls = 0
    const tracker = new SelectionTracker(async () => {
      const r = reads[Math.min(calls, reads.length - 1)] ?? read(null)
      calls += 1
      await Promise.resolve()
      return r
    })
    await Promise.all([tracker.refresh(), tracker.refresh(), tracker.refresh()])
    expect(calls).toBe(2)
    expect(tracker.current.order).toEqual(['a', 'b'])
  })

  it('survives a read that fails', async () => {
    const tracker = new SelectionTracker(async () => {
      throw new Error('slide is changing')
    })
    await expect(tracker.refresh()).resolves.toBeUndefined()
  })
})

describe('tracking in PowerPoint', () => {
  it('records the order shapes join the selection, which PowerPoint does not report', async () => {
    const deck = createPresentation([slide([shape({ id: 'a' }), shape({ id: 'b' }), shape({ id: 'c' })], { id: 's1' })])
    installFakePowerPoint(deck)
    await selectionTracker.start()
    expect(deck.handlers).toHaveLength(1)
    await clickShapes(deck, ['c'])
    await clickShapes(deck, ['a', 'c'])
    await clickShapes(deck, ['a', 'b', 'c'])
    expect(selectionTracker.current.order).toEqual(['c', 'a', 'b'])
    expect(anchorHint('s1', ['a', 'b', 'c'])).toEqual({ anchorId: 'c', explicit: false, order: ['c', 'a', 'b'] })
    await selectionTracker.stop()
    expect(deck.handlers).toHaveLength(0)
  })

  it('does not listen in a PowerPoint without selection events (PowerPointApi 1.5)', async () => {
    const deck = createPresentation([slide([shape({ id: 'a' })], { id: 's1' })], { apiVersion: '1.4' })
    installFakePowerPoint(deck)
    await selectionTracker.start()
    expect(deck.handlers).toHaveLength(0)
    expect(selectionTracker.listening).toBe(false)
  })

  it('sets the one selected shape as the anchor, and clears it', async () => {
    const deck = createPresentation([slide([shape({ id: 'a', name: 'Box A' }), shape({ id: 'b' })], { id: 's1' })])
    installFakePowerPoint(deck)
    await clickShapes(deck, ['a'])
    expect(await selectionTracker.setAnchorFromSelection()).toBe('“Box A” is the anchor. Alignment to the anchor and Same size use it.')
    expect(selectionTracker.current.anchor).toEqual({ id: 'a', name: 'Box A', slideId: 's1' })
    await clickShapes(deck, ['a', 'b'])
    await expect(selectionTracker.setAnchorFromSelection()).rejects.toThrow('Select one shape to make it the anchor.')
    selectionTracker.clearAnchor()
    expect(selectionTracker.current.anchor).toBeNull()
  })
})

describe('the anchor hint', () => {
  const state = (over: Partial<SelectionState>): SelectionState => ({ ...emptyState(), at: 1, ...over })

  it('prefers the anchor set by hand on the same slide, even outside the selection', () => {
    const s = state({ slideId: 's1', order: ['b', 'a'], anchor: { id: 'z', name: 'Z', slideId: 's1' } })
    expect(anchorHint('s1', ['a', 'b'], s)).toEqual({ anchorId: 'z', explicit: true, order: ['b', 'a'] })
  })

  it('falls back to the first selected shape, then to none', () => {
    const s = state({ slideId: 's1', order: ['b', 'a'], anchor: { id: 'z', name: 'Z', slideId: 's2' } })
    expect(anchorHint('s1', ['a', 'b'], s)).toEqual({ anchorId: 'b', explicit: false, order: ['b', 'a'] })
    expect(anchorHint('s1', ['a'], s).anchorId).toBe('a')
    expect(anchorHint('s3', ['a', 'b'], s)).toEqual({ anchorId: null, explicit: false, order: [] })
    expect(anchorHint('s1', ['x'], s).anchorId).toBeNull()
  })
})

describe('the saved state, for the ribbon', () => {
  it('saves every change and reads it back while it is fresh', () => {
    vi.stubGlobal('localStorage', memoryStorage())
    let now = 1_000_000
    const tracker = new SelectionTracker(async () => read(null), () => now)
    tracker.apply(read('s1', 'a', 'b'))
    expect(loadState(now)).toMatchObject({ slideId: 's1', order: ['a', 'b'], at: now })
    now += 13 * 60 * 60 * 1000
    expect(loadState(now)).toEqual(emptyState())
  })

  it('ignores storage that is missing, broken or refuses access', () => {
    expect(loadState()).toEqual(emptyState())
    const broken = memoryStorage()
    broken.setItem('bwb.ppt.selection.v1', '{not json')
    vi.stubGlobal('localStorage', broken)
    expect(loadState()).toEqual(emptyState())
    vi.stubGlobal('localStorage', { getItem: () => { throw new Error('denied') }, setItem: () => { throw new Error('denied') } })
    expect(loadState()).toEqual(emptyState())
    const tracker = new SelectionTracker(async () => read(null))
    expect(() => tracker.apply(read('s1', 'a'))).not.toThrow()
  })
})
