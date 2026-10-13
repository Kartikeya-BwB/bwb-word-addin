import { describe, expect, it } from 'vitest'
import { align, distribute, equaliseGaps, stack } from './align.ts'
import type { AlignEdge } from './align.ts'
import { visualBounds } from './bounds.ts'
import type { Box } from './types.ts'

const box = (id: string, left: number, top: number, width: number, height: number, rotation?: number): Box =>
  rotation === undefined ? { id, left, top, width, height } : { id, left, top, width, height, rotation }

const lefts = (boxes: Box[]): number[] => boxes.map((b) => b.left)
const tops = (boxes: Box[]): number[] => boxes.map((b) => b.top)

describe('align', () => {
  const a = box('a', 0, 0, 100, 50)
  const b = box('b', 200, 100, 50, 50)
  const c = box('c', 50, 300, 80, 40)

  it('does nothing with nothing selected', () => {
    expect(align([], 'left', { reference: 'selection' })).toEqual([])
  })

  it.each<[AlignEdge, number[], number[]]>([
    ['left', [0, 0, 0], [0, 100, 300]],
    ['centre', [75, 100, 85], [0, 100, 300]],
    ['right', [150, 200, 170], [0, 100, 300]],
    ['top', [0, 200, 50], [0, 0, 0]],
    ['middle', [0, 200, 50], [145, 145, 150]],
    ['bottom', [0, 200, 50], [290, 290, 300]],
  ])('lines up %s edges with the outermost one in the selection', (edge, xs, ys) => {
    const out = align([a, b, c], edge, { reference: 'selection' })
    expect(lefts(out)).toEqual(xs)
    expect(tops(out)).toEqual(ys)
    expect(out.map((o) => [o.id, o.width, o.height])).toEqual([['a', 100, 50], ['b', 50, 50], ['c', 80, 40]])
  })

  it('lines one shape up with the margins when the selection is all there is', () => {
    const x = box('x', 300, 300, 100, 50)
    expect(align([x], 'left', { reference: 'selection' })[0]?.left).toBe(40)
    expect(align([x], 'right', { reference: 'selection' })[0]?.left).toBe(820)
    expect(align([x], 'top', { reference: 'selection' })[0]?.top).toBe(28)
    expect(align([x], 'bottom', { reference: 'selection' })[0]?.top).toBe(454)
  })

  it('keeps the anchor still and lines the others up with it', () => {
    const out = align([a, b, c], 'left', { reference: 'anchor', anchor: b })
    expect(lefts(out)).toEqual([200, 200, 200])
    expect(out[1]).toEqual(b)
  })

  it('uses the largest shape as the anchor when none is set', () => {
    const big = box('big', 500, 0, 300, 300)
    expect(lefts(align([a, big], 'right', { reference: 'anchor' }))).toEqual([700, 500])
  })

  it('lines up with an anchor outside the selection', () => {
    const anchor = box('z', 600, 400, 100, 100)
    const out = align([a, b], 'bottom', { reference: 'anchor', anchor })
    expect(tops(out)).toEqual([450, 450])
  })

  it('lines up with the slide, the margins and the content area', () => {
    const x = box('x', 300, 300, 100, 50)
    expect(align([x], 'centre', { reference: 'slide' })[0]?.left).toBe(430)
    expect(align([x], 'middle', { reference: 'slide' })[0]?.top).toBe(245)
    expect(align([x], 'top', { reference: 'content' })[0]?.top).toBe(108)
    expect(align([x], 'bottom', { reference: 'content' })[0]?.top).toBe(418)
    expect(align([x], 'right', { reference: 'margins' })[0]?.left).toBe(820)
    expect(align([x], 'centre', { reference: 'slide', slide: { width: 720, height: 540 } })[0]?.left).toBe(310)
  })

  it('moves the selection as a group, keeping its layout', () => {
    const out = align([a, b], 'centre', { reference: 'slide', asGroup: true })
    expect(lefts(out)).toEqual([355, 555])
    expect(tops(align([a, b], 'top', { reference: 'content', asGroup: true }))).toEqual([108, 208])
    expect(align([a, b], 'left', { reference: 'selection', asGroup: true })).toEqual([a, b])
  })

  it('moves the others as a group to the anchor, leaving the anchor', () => {
    const out = align([a, b, c], 'left', { reference: 'anchor', anchor: b, asGroup: true })
    expect(lefts(out)).toEqual([200, 200, 250])
  })

  it('lines rotated shapes up by what people see, keeping size and rotation', () => {
    const r = box('r', 0, 0, 100, 20, 90)
    const anchor = box('a', 200, 0, 50, 50)
    const [moved] = align([r], 'left', { reference: 'anchor', anchor })
    expect(moved).toEqual({ ...r, left: 160 })
    expect(moved && visualBounds(moved).left).toBe(200)
    const [top] = align([r], 'top', { reference: 'anchor', anchor })
    expect(top && visualBounds(top).top).toBe(0)
  })

  it('handles shapes off the slide and overlapping shapes', () => {
    const off = box('off', -50, -20, 30, 30)
    expect(lefts(align([off, a], 'left', { reference: 'selection' }))).toEqual([-50, -50])
    const overlap = box('o', 10, 10, 100, 50)
    expect(lefts(align([a, overlap], 'centre', { reference: 'selection' }))).toEqual([5, 5])
  })
})

describe('distribute', () => {
  it('needs three shapes to space them between the outer two', () => {
    expect(distribute([], { axis: 'horizontal' })).toEqual([])
    const two = [box('a', 0, 0, 10, 10), box('b', 100, 0, 10, 10)]
    expect(distribute(two, { axis: 'horizontal' })).toEqual(two)
    expect(distribute(two.slice(0, 1), { axis: 'horizontal' })).toEqual(two.slice(0, 1))
  })

  it('makes equal gaps between shapes of different sizes, keeping the outer two', () => {
    const out = distribute([box('a', 0, 0, 20, 10), box('b', 100, 7, 60, 10), box('c', 300, 0, 100, 10)], { axis: 'horizontal' })
    expect(lefts(out)).toEqual([0, 130, 300])
    expect(tops(out)).toEqual([0, 7, 0])
  })

  it('returns shapes in the order given, whatever their order on the slide', () => {
    const out = distribute([box('c', 400, 0, 100, 50), box('a', 0, 0, 100, 50), box('b', 150, 0, 50, 50)], { axis: 'horizontal' })
    expect(out.map((o) => [o.id, o.left])).toEqual([['c', 400], ['a', 0], ['b', 225]])
  })

  it('works top to bottom', () => {
    const out = distribute([box('a', 0, 0, 10, 20), box('b', 0, 100, 10, 60), box('c', 0, 300, 10, 100)], { axis: 'vertical' })
    expect(tops(out)).toEqual([0, 130, 300])
  })

  it('makes equal distances between centres', () => {
    const out = distribute([box('a', 0, 0, 20, 10), box('b', 100, 0, 60, 10), box('c', 300, 0, 100, 10)], { axis: 'horizontal', spacing: 'centres' })
    expect(lefts(out)).toEqual([0, 150, 300])
  })

  it('spreads shapes across the margins, the content area or the slide', () => {
    const two = [box('a', 0, 0, 100, 50), box('b', 300, 0, 100, 50)]
    expect(lefts(distribute(two, { axis: 'horizontal', within: 'margins' }))).toEqual([40, 820])
    expect(lefts(distribute(two, { axis: 'horizontal', within: 'slide' }))).toEqual([0, 860])
    const tall = [box('a', 0, 0, 10, 100), box('b', 0, 150, 10, 100)]
    expect(tops(distribute(tall, { axis: 'vertical', within: 'content' }))).toEqual([108, 368])
    expect(lefts(distribute(two, { axis: 'horizontal', within: 'margins', spacing: 'centres' }))).toEqual([40, 820])
  })

  it('centres a single shape in an area', () => {
    expect(lefts(distribute([box('a', 0, 0, 100, 50)], { axis: 'horizontal', within: 'margins' }))).toEqual([430])
    expect(tops(distribute([box('a', 0, 0, 100, 60)], { axis: 'vertical', within: 'content' }))).toEqual([258])
  })

  it('stacks with an exact gap from the first shape, whatever the extent', () => {
    const out = distribute([box('a', 0, 0, 100, 10), box('b', 300, 0, 50, 10), box('c', 500, 0, 100, 10)], { axis: 'horizontal', gap: 10, within: 'margins' })
    expect(lefts(out)).toEqual([0, 110, 170])
  })

  it('spaces rotated shapes by their bounds', () => {
    const out = distribute([box('r', 0, 0, 100, 20, 90), box('a', 100, 0, 20, 20), box('c', 300, 0, 20, 20)], { axis: 'horizontal' })
    expect(lefts(out)).toEqual([0, 170, 300])
  })

  it('overlaps evenly when the shapes do not fit', () => {
    const out = distribute([box('a', 0, 0, 100, 10), box('b', 50, 0, 100, 10), box('c', 60, 0, 100, 10)], { axis: 'horizontal' })
    expect(lefts(out)).toEqual([0, 30, 60])
  })

  it('handles negative positions', () => {
    const out = distribute([box('a', -200, 0, 50, 10), box('b', -10, 0, 50, 10), box('c', 100, 0, 50, 10)], { axis: 'horizontal' })
    expect(lefts(out)).toEqual([-200, -50, 100])
  })
})

describe('equal gaps and stacking', () => {
  it('sets every gap to the average gap, keeping the first shape and the order', () => {
    const out = equaliseGaps([box('a', 0, 0, 100, 10), box('b', 150, 0, 50, 10), box('c', 210, 0, 100, 10)], 'horizontal')
    expect(lefts(out)).toEqual([0, 130, 210])
    const down = equaliseGaps([box('a', 0, 0, 10, 10), box('b', 0, 30, 10, 10), box('c', 0, 100, 10, 10)], 'vertical')
    expect(tops(down)).toEqual([0, 50, 100])
  })

  it('leaves fewer than three shapes alone', () => {
    const two = [box('a', 0, 0, 10, 10), box('b', 100, 0, 10, 10)]
    expect(equaliseGaps(two, 'horizontal')).toEqual(two)
  })

  it('butts shapes together edge to edge, or with a gap', () => {
    const row = [box('a', 0, 0, 100, 10), box('b', 300, 5, 50, 10), box('c', 500, 0, 100, 10)]
    expect(lefts(stack(row, 'horizontal'))).toEqual([0, 100, 150])
    expect(lefts(stack(row, 'horizontal', 8))).toEqual([0, 108, 166])
    expect(tops(stack(row, 'horizontal'))).toEqual([0, 5, 0])
    const column = [box('a', 0, 0, 10, 30), box('b', 0, 200, 10, 20)]
    expect(tops(stack(column, 'vertical', 4))).toEqual([0, 34])
    expect(stack(row.slice(0, 1), 'horizontal')).toEqual(row.slice(0, 1))
  })
})
