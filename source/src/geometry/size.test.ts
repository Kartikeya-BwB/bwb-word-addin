import { describe, expect, it } from 'vitest'
import { columnEdges } from '../brand/slideGrid.ts'
import { visualBounds } from './bounds.ts'
import { matchSize, setFrame, snapToGrid, stretch } from './size.ts'
import type { Box } from './types.ts'

const box = (id: string, left: number, top: number, width: number, height: number, rotation?: number): Box =>
  rotation === undefined ? { id, left, top, width, height } : { id, left, top, width, height, rotation }

const sizes = (boxes: Box[]): Array<[number, number]> => boxes.map((b) => [b.width, b.height])

describe('same size', () => {
  const a = box('a', 0, 0, 100, 50)
  const b = box('b', 200, 0, 60, 30)
  const c = box('c', 300, 0, 200, 10)

  it('does nothing with nothing selected', () => {
    expect(matchSize([], { dimension: 'width', source: 'anchor' })).toEqual([])
  })

  it('gives every shape the anchor’s width, height or both', () => {
    expect(sizes(matchSize([a, b, c], { dimension: 'width', source: 'anchor', anchor: b }))).toEqual([[60, 50], [60, 30], [60, 10]])
    expect(sizes(matchSize([a, b, c], { dimension: 'height', source: 'anchor', anchor: b }))).toEqual([[100, 30], [60, 30], [200, 30]])
    expect(sizes(matchSize([a, b, c], { dimension: 'both', source: 'anchor', anchor: b }))).toEqual([[60, 30], [60, 30], [60, 30]])
  })

  it('keeps each shape’s top-left corner', () => {
    const out = matchSize([a, b, c], { dimension: 'both', source: 'anchor', anchor: b })
    expect(out.map((o) => [o.left, o.top])).toEqual([[0, 0], [200, 0], [300, 0]])
  })

  it('takes the largest shape as the anchor when none is set, and leaves a lone shape alone', () => {
    expect(sizes(matchSize([b, a, c], { dimension: 'width', source: 'anchor' }))).toEqual([[100, 30], [100, 50], [100, 10]])
    expect(matchSize([a], { dimension: 'both', source: 'anchor' })).toEqual([a])
  })

  it('matches an anchor outside the selection', () => {
    const anchor = box('z', 0, 0, 77, 1)
    expect(sizes(matchSize([a, b], { dimension: 'width', source: 'anchor', anchor }))).toEqual([[77, 50], [77, 30]])
  })

  it('matches the first shape given', () => {
    expect(sizes(matchSize([c, a], { dimension: 'height', source: 'first' }))).toEqual([[200, 10], [100, 10]])
  })

  it('takes the largest or smallest width or height on its own', () => {
    expect(sizes(matchSize([a, b, c], { dimension: 'width', source: 'largest' }))).toEqual([[200, 50], [200, 30], [200, 10]])
    expect(sizes(matchSize([a, b, c], { dimension: 'height', source: 'smallest' }))).toEqual([[100, 10], [60, 10], [200, 10]])
  })

  it('takes the size of the largest or smallest shape for both', () => {
    expect(sizes(matchSize([a, b, c], { dimension: 'both', source: 'largest' }))).toEqual([[100, 50], [100, 50], [100, 50]])
    expect(sizes(matchSize([a, b, c], { dimension: 'both', source: 'smallest' }))).toEqual([[60, 30], [60, 30], [60, 30]])
  })

  it('keeps proportions when asked', () => {
    expect(sizes(matchSize([a, b, c], { dimension: 'width', source: 'anchor', anchor: b, keepAspect: true }))).toEqual([[60, 30], [60, 30], [60, 3]])
    expect(sizes(matchSize([a, b, c], { dimension: 'height', source: 'anchor', anchor: b, keepAspect: true }))).toEqual([[60, 30], [60, 30], [600, 30]])
  })

  it('leaves the other side of a flat line alone when keeping proportions', () => {
    const line = box('l', 0, 0, 100, 0)
    expect(sizes(matchSize([line], { dimension: 'height', source: 'anchor', anchor: b, keepAspect: true }))).toEqual([[100, 30]])
  })

  it('reads a quarter-turned shape’s width as its frame’s height', () => {
    const r = box('r', 0, 0, 100, 20, 90)
    const [out] = matchSize([r], { dimension: 'width', source: 'anchor', anchor: b })
    expect(out).toMatchObject({ width: 100, height: 60, rotation: 90 })
    expect(out && visualBounds(out)).toEqual({ left: 40, top: -40, width: 60, height: 100 })
  })
})

describe('exact frame', () => {
  it('sets the given values and keeps the rest', () => {
    expect(setFrame([box('a', 0, 0, 100, 50)], { left: 10.005, width: 20 })).toEqual([box('a', 10.01, 0, 20, 50)])
    expect(setFrame([box('a', 0, 0, 100, 50), box('b', 5, 5, 5, 5)], { top: 108 }).map((b) => b.top)).toEqual([108, 108])
    expect(setFrame([box('a', -5, 0, 100, 50)], {})).toEqual([box('a', -5, 0, 100, 50)])
  })

  it('refuses a size of zero or less', () => {
    expect(() => setFrame([box('a', 0, 0, 1, 1)], { width: 0 })).toThrow(/more than zero/)
    expect(() => setFrame([box('a', 0, 0, 1, 1)], { height: -3 })).toThrow(/more than zero/)
  })
})

describe('stretch', () => {
  const anchor = box('b', 200, 0, 60, 30)

  it('moves an edge to the anchor’s edge, keeping the opposite edge and the anchor', () => {
    const out = stretch([box('a', 0, 0, 100, 50), anchor], 'right', { reference: 'anchor', anchor })
    expect(out).toEqual([box('a', 0, 0, 260, 50), anchor])
  })

  it('leaves a shape whose edge would pass its other edge', () => {
    const c = box('c', 300, 0, 200, 10)
    expect(stretch([c], 'right', { reference: 'anchor', anchor })).toEqual([c])
  })

  it('stretches to the margins, the content area and the slide', () => {
    expect(stretch([box('a', 100, 0, 100, 50)], 'left', { reference: 'margins' })).toEqual([box('a', 40, 0, 160, 50)])
    expect(stretch([box('a', 0, 200, 100, 50)], 'top', { reference: 'content' })).toEqual([box('a', 0, 108, 100, 142)])
    expect(stretch([box('a', 0, 200, 100, 50)], 'bottom', { reference: 'slide' })).toEqual([box('a', 0, 200, 100, 340)])
    expect(stretch([box('a', 0, 200, 100, 50)], 'right', { reference: 'margins' })).toEqual([box('a', 0, 200, 920, 50)])
  })

  it('stretches to the outermost edge of the selection', () => {
    const out = stretch([box('a', 0, 0, 100, 50), anchor], 'right', { reference: 'selection' })
    expect(out.map((o) => o.width)).toEqual([260, 60])
  })

  it('stretches a rotated shape by what people see', () => {
    const [r] = stretch([box('r', 0, 0, 100, 20, 90)], 'right', { reference: 'margins' })
    expect(r && visualBounds(r)).toEqual({ left: 40, top: -40, width: 880, height: 100 })
  })

  it('treats the only shape as its own anchor', () => {
    const a = box('a', 0, 0, 10, 10)
    expect(stretch([a], 'left', { reference: 'anchor' })).toEqual([a])
    expect(stretch([], 'left', { reference: 'anchor' })).toEqual([])
  })
})

describe('snap to the BwB grid', () => {
  const cols = columnEdges()
  const col = (i: number): { left: number; right: number } => cols[i] ?? { left: NaN, right: NaN }

  it('snaps the left edge to a column start and the right edge to a column end', () => {
    const [out] = snapToGrid([box('a', 45, 7, 200, 50)])
    expect(out?.left).toBeCloseTo(col(0).left, 2)
    expect((out?.left ?? 0) + (out?.width ?? 0)).toBeCloseTo(col(2).right, 2)
    expect(out?.top).toBe(7)
    expect(out?.height).toBe(50)
  })

  it('gives a shape narrower than a column one column', () => {
    const [out] = snapToGrid([box('a', 120, 0, 10, 10)])
    expect(out?.left).toBeCloseTo(col(1).left, 2)
    expect(out?.width).toBeCloseTo(col(1).right - col(1).left, 2)
  })

  it('brings shapes off the slide onto the grid', () => {
    const [left] = snapToGrid([box('a', -100, 0, 50, 10)])
    expect(left?.left).toBeCloseTo(col(0).left, 2)
    const [right] = snapToGrid([box('b', 900, 0, 100, 10)])
    expect(right?.left).toBeCloseTo(col(11).left, 2)
    expect((right?.left ?? 0) + (right?.width ?? 0)).toBeCloseTo(col(11).right, 2)
  })

  it('snaps the bounds of a rotated shape', () => {
    const [r] = snapToGrid([box('r', 0, 0, 100, 20, 90)])
    const v = r && visualBounds(r)
    expect(v?.left).toBeCloseTo(col(0).left, 2)
    expect(v?.width).toBeCloseTo(col(0).right - col(0).left, 2)
  })

  it('leaves a shape already on the grid where it is', () => {
    const on = box('a', col(3).left, 0, col(5).right - col(3).left, 10)
    const [out] = snapToGrid([on])
    expect(out?.left).toBeCloseTo(on.left, 2)
    expect(out?.width).toBeCloseTo(on.width, 2)
    expect(snapToGrid([])).toEqual([])
  })
})
