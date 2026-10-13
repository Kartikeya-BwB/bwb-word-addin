// Laying shapes out: grids, tidying a rough layout into rows and columns, swapping and nudging.

import { GUTTER } from '../brand/slideGrid.ts'
import { centreX, centreY, copy, effectiveSize, moveBoundsTo, resizeEffective, translate, union, visualBounds } from './bounds.ts'
import type { Box, Rect } from './types.ts'
import { GeometryError } from './types.ts'

const median = (values: readonly number[]): number => {
  const s = [...values].sort((a, b) => a - b)
  const mid = Math.floor(s.length / 2)
  return s.length % 2 ? (s[mid] ?? 0) : ((s[mid - 1] ?? 0) + (s[mid] ?? 0)) / 2
}

const mean = (values: readonly number[]): number => (values.length ? values.reduce((a, b) => a + b, 0) / values.length : 0)

interface Item {
  box: Box
  index: number
  v: Rect
}

const itemsOf = (boxes: readonly Box[]): Item[] => boxes.map((box, index) => ({ box, index, v: visualBounds(box) }))

/** Groups items whose key lies within `tolerance` of the previous item's key, in key order. */
function cluster(items: readonly Item[], key: (i: Item) => number, tolerance: number): Item[][] {
  const out: Item[][] = []
  let previous = -Infinity
  for (const it of [...items].sort((a, b) => key(a) - key(b) || a.index - b.index)) {
    const k = key(it)
    const current = out[out.length - 1]
    if (current && k - previous <= tolerance) current.push(it)
    else out.push([it])
    previous = k
  }
  return out
}

/** Rows by vertical centre; a row's shapes are within half the typical shape height of each other. */
function rowsOf(items: readonly Item[], tolerance?: number): Item[][] {
  return cluster(items, (i) => centreY(i.v), tolerance ?? Math.max(4, median(items.map((i) => i.v.height)) / 2))
}

function columnsOf(items: readonly Item[], tolerance?: number): Item[][] {
  return cluster(items, (i) => centreX(i.v), tolerance ?? Math.max(4, median(items.map((i) => i.v.width)) / 2))
}

/** Boxes in reading order: rows from the top, each row from the left. */
export function readingOrder(boxes: readonly Box[]): Box[] {
  return rowsOf(itemsOf(boxes)).flatMap((row) => row.sort((a, b) => a.v.left - b.v.left || a.index - b.index).map((i) => i.box))
}

export interface GridOptions {
  columns: number
  /** Gaps in points; the BwB gutter by default. */
  gapX?: number
  gapY?: number
  /** Uniform cells give every shape the same size; 'keep' (the default) keeps sizes, columns as wide as their widest. */
  cells?: 'uniform' | 'keep'
  /** Uniform cells only: fill this area. Otherwise cells take the first shape's size. */
  area?: Rect
}

/** Tiles shapes into a grid of n columns, in reading order, from the top-left corner of the selection (or area). */
export function grid(boxes: readonly Box[], options: GridOptions): Box[] {
  if (!boxes.length) return []
  if (!(options.columns >= 1)) throw new GeometryError('A grid needs one column or more.')
  const ordered = readingOrder(boxes)
  const cols = Math.min(Math.floor(options.columns), ordered.length)
  const rows = Math.ceil(ordered.length / cols)
  const gapX = options.gapX ?? GUTTER
  const gapY = options.gapY ?? GUTTER
  const origin = options.area ?? union(ordered.map(visualBounds))
  if (!origin) return boxes.map(copy)
  const placed = new Map<Box, Box>()
  if (options.cells === 'uniform') {
    const first = ordered[0]
    const size = options.area
      ? { width: (options.area.width - (cols - 1) * gapX) / cols, height: (options.area.height - (rows - 1) * gapY) / rows }
      : first ? effectiveSize(first) : { width: 0, height: 0 }
    if (!(size.width > 0) || !(size.height > 0)) throw new GeometryError('The area is too small for that grid.')
    ordered.forEach((b, k) => {
      const sized = resizeEffective(b, size.width, size.height)
      placed.set(b, moveBoundsTo(sized, origin.left + (k % cols) * (size.width + gapX), origin.top + Math.floor(k / cols) * (size.height + gapY)))
    })
  } else {
    const vs = ordered.map(visualBounds)
    const widths = Array.from({ length: cols }, (_, c) => Math.max(...vs.filter((_, k) => k % cols === c).map((v) => v.width)))
    const heights = Array.from({ length: rows }, (_, r) => Math.max(...vs.filter((_, k) => Math.floor(k / cols) === r).map((v) => v.height)))
    const offsets = (sizes: number[], gap: number): number[] => sizes.map((_, i) => sizes.slice(0, i).reduce((sum, s) => sum + s + gap, 0))
    const xs = offsets(widths, gapX)
    const ys = offsets(heights, gapY)
    ordered.forEach((b, k) => placed.set(b, moveBoundsTo(b, origin.left + (xs[k % cols] ?? 0), origin.top + (ys[Math.floor(k / cols)] ?? 0))))
  }
  return boxes.map((b) => placed.get(b) ?? copy(b))
}

/** Two shapes swap places: each takes the other's top-left corner and keeps its own size. */
export function swap(boxes: readonly Box[]): Box[] {
  const [a, b] = boxes
  if (boxes.length !== 2 || !a || !b) throw new GeometryError('Select two shapes to swap.')
  const va = visualBounds(a)
  const vb = visualBounds(b)
  return [translate(a, vb.left - va.left, vb.top - va.top), translate(b, va.left - vb.left, va.top - vb.top)]
}

export interface TidyOptions {
  /** How far apart, in points, centres can be and still share a row or column. Half the typical size by default. */
  tolerance?: number
  /** Gaps in points; by default the average of the gaps there are now. */
  gapX?: number
  gapY?: number
}

/**
 * Finds the rows and columns in a rough layout and lines them up: every shape in a row gets the row's top edge,
 * every shape in a column the column's left edge, and rows and columns are spaced with equal gaps, starting from
 * the top-left shape.
 */
export function tidyUp(boxes: readonly Box[], options: TidyOptions = {}): Box[] {
  if (boxes.length < 2) return boxes.map(copy)
  const items = itemsOf(boxes)
  const lines = (groups: Item[][], start: (r: Rect) => number, size: (r: Rect) => number, gapOption: number | undefined): Map<Item, number> => {
    const lo = groups.map((g) => Math.min(...g.map((i) => start(i.v))))
    const hi = groups.map((g) => Math.max(...g.map((i) => start(i.v) + size(i.v))))
    const extents = groups.map((g) => Math.max(...g.map((i) => size(i.v))))
    const gap = gapOption ?? Math.max(0, mean(lo.slice(1).map((l, k) => l - (hi[k] ?? l))))
    const at = new Map<Item, number>()
    let pos = lo[0] ?? 0
    groups.forEach((g, k) => {
      g.forEach((i) => at.set(i, pos))
      pos += (extents[k] ?? 0) + gap
    })
    return at
  }
  const tops = lines(rowsOf(items, options.tolerance), (r) => r.top, (r) => r.height, options.gapY)
  const lefts = lines(columnsOf(items, options.tolerance), (r) => r.left, (r) => r.width, options.gapX)
  return items.map((i) => translate(i.box, (lefts.get(i) ?? i.v.left) - i.v.left, (tops.get(i) ?? i.v.top) - i.v.top))
}

/** Moves every shape by the same amount, in points. */
export function nudge(boxes: readonly Box[], dx: number, dy: number): Box[] {
  return boxes.map((b) => translate(b, dx, dy))
}
