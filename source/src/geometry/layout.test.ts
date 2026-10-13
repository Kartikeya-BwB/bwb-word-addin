import { describe, expect, it } from 'vitest'
import { visualBounds } from './bounds.ts'
import { grid, nudge, readingOrder, swap, tidyUp } from './layout.ts'
import type { Box } from './types.ts'

const box = (id: string, left: number, top: number, width: number, height: number, rotation?: number): Box =>
  rotation === undefined ? { id, left, top, width, height } : { id, left, top, width, height, rotation }

const corners = (boxes: Box[]): Array<[string, number, number]> => boxes.map((b) => [b.id, b.left, b.top])

describe('reading order', () => {
  it('reads rows from the top, each from the left, allowing for a rough layout', () => {
    const order = readingOrder([box('d', 200, 100, 50, 50), box('a', 0, 0, 50, 50), box('c', 0, 100, 50, 50), box('b', 200, 5, 50, 50)])
    expect(order.map((b) => b.id)).toEqual(['a', 'b', 'c', 'd'])
  })
})

describe('grid', () => {
  const row = [box('a', 0, 0, 50, 50), box('b', 100, 0, 50, 50), box('c', 200, 0, 50, 50), box('d', 300, 0, 50, 50)]

  it('tiles shapes into columns from the top-left of the selection', () => {
    expect(corners(grid(row, { columns: 2, gapX: 10, gapY: 10 }))).toEqual([['a', 0, 0], ['b', 60, 0], ['c', 0, 60], ['d', 60, 60]])
  })

  it('uses the BwB gutter when no gaps are given', () => {
    expect(corners(grid(row, { columns: 2 }))).toEqual([['a', 0, 0], ['b', 66, 0], ['c', 0, 66], ['d', 66, 66]])
  })

  it('keeps sizes, making each column as wide as its widest shape and each row as tall as its tallest', () => {
    const mixed = [box('a', 0, 0, 50, 50), box('b', 100, 0, 80, 20), box('c', 200, 0, 30, 70), box('d', 300, 0, 50, 50)]
    const out = grid(mixed, { columns: 2, gapX: 10, gapY: 10 })
    expect(corners(out)).toEqual([['a', 0, 0], ['b', 60, 0], ['c', 0, 60], ['d', 60, 60]])
    expect(out.map((b) => b.width)).toEqual([50, 80, 30, 50])
  })

  it('gives every shape the first shape’s size in uniform cells', () => {
    const mixed = [box('a', 0, 0, 50, 40), box('b', 100, 0, 80, 20), box('c', 200, 0, 30, 70)]
    const out = grid(mixed, { columns: 3, gapX: 5, gapY: 5, cells: 'uniform' })
    expect(out.map((b) => [b.left, b.width, b.height])).toEqual([[0, 50, 40], [55, 50, 40], [110, 50, 40]])
  })

  it('fills an area with uniform cells', () => {
    const six = Array.from({ length: 6 }, (_, i) => box(`s${i}`, i * 10, 0, 10, 10))
    const out = grid(six, { columns: 3, cells: 'uniform', area: { left: 40, top: 108, width: 880, height: 360 } })
    expect(out[0]).toMatchObject({ left: 40, top: 108, width: 282.67, height: 172 })
    expect(out[4]?.left).toBeCloseTo(40 + 298.67, 1)
    expect(out[4]?.top).toBe(296)
  })

  it('keeps the order given, places by reading order, and handles one shape or too many columns', () => {
    const shuffled = [row[3], row[0], row[2], row[1]].filter((b): b is Box => b !== undefined)
    expect(corners(grid(shuffled, { columns: 2, gapX: 10, gapY: 10 }))).toEqual([['d', 60, 60], ['a', 0, 0], ['c', 0, 60], ['b', 60, 0]])
    expect(grid(row.slice(0, 1), { columns: 3 })).toEqual(row.slice(0, 1))
    expect(corners(grid(row.slice(0, 2), { columns: 9, gapX: 0 }))).toEqual([['a', 0, 0], ['b', 50, 0]])
    expect(grid([], { columns: 2 })).toEqual([])
  })

  it('refuses fewer than one column and an area too small', () => {
    expect(() => grid(row, { columns: 0 })).toThrow(/one column/)
    expect(() => grid(row, { columns: 2, cells: 'uniform', area: { left: 0, top: 0, width: 10, height: 10 }, gapX: 20 })).toThrow(/too small/)
  })

  it('places rotated shapes by their bounds', () => {
    const out = grid([box('r', 0, 0, 100, 20, 90), box('b', 100, 0, 20, 20)], { columns: 2, gapX: 10 })
    const r = out[0]
    expect(r && visualBounds(r)).toMatchObject({ left: 40, top: -40 })
    expect(out[1]).toMatchObject({ left: 70, top: -40 })
  })
})

describe('swap', () => {
  it('swaps the top-left corners of two shapes, keeping their sizes', () => {
    const out = swap([box('a', 0, 0, 100, 50), box('b', 300, 200, 40, 40)])
    expect(out).toEqual([box('a', 300, 200, 100, 50), box('b', 0, 0, 40, 40)])
  })

  it('swaps rotated shapes by their bounds', () => {
    const [r, b] = swap([box('r', 0, 0, 100, 20, 90), box('b', 300, 200, 40, 40)])
    expect(r && visualBounds(r)).toMatchObject({ left: 300, top: 200 })
    expect(b).toMatchObject({ left: 40, top: -40 })
  })

  it('needs exactly two shapes', () => {
    expect(() => swap([])).toThrow('Select two shapes to swap.')
    expect(() => swap([box('a', 0, 0, 1, 1)])).toThrow('Select two shapes to swap.')
    expect(() => swap([box('a', 0, 0, 1, 1), box('b', 0, 0, 1, 1), box('c', 0, 0, 1, 1)])).toThrow('Select two shapes to swap.')
  })
})

describe('tidy up', () => {
  it('finds rows and columns in a rough layout and spaces them evenly', () => {
    const rough = [box('a', 0, 0, 100, 50), box('b', 155, 4, 100, 50), box('c', 3, 98, 100, 50), box('d', 150, 103, 100, 50)]
    expect(corners(tidyUp(rough))).toEqual([['a', 0, 0], ['b', 147, 0], ['c', 0, 94], ['d', 147, 94]])
  })

  it('lines up a single row and evens out its gaps', () => {
    const out = tidyUp([box('a', 0, 0, 100, 50), box('b', 130, 5, 100, 50), box('c', 300, -3, 100, 50)])
    expect(corners(out)).toEqual([['a', 0, -3], ['b', 150, -3], ['c', 300, -3]])
  })

  it('takes exact gaps and a tolerance', () => {
    const rough = [box('a', 0, 0, 100, 50), box('b', 155, 4, 100, 50)]
    expect(corners(tidyUp(rough, { gapX: 16 }))).toEqual([['a', 0, 0], ['b', 116, 0]])
    expect(corners(tidyUp(rough, { tolerance: 1, gapY: 0 }))).toEqual([['a', 0, 0], ['b', 155, 50]])
  })

  it('leaves fewer than two shapes alone', () => {
    expect(tidyUp([])).toEqual([])
    expect(tidyUp([box('a', 3, 4, 5, 6)])).toEqual([box('a', 3, 4, 5, 6)])
  })
})

describe('nudge', () => {
  it('moves every shape by the same amount', () => {
    expect(corners(nudge([box('a', 0, 0, 1, 1), box('b', -5, 5, 1, 1)], 2.835, -1))).toEqual([['a', 2.84, -1], ['b', -2.16, 4]])
    expect(nudge([], 1, 1)).toEqual([])
  })
})
