// Lining shapes up and spacing them out: align, distribute, equal gaps and stacking. Positions use each shape's
// rotated bounds, so a turned shape lines up by what people see; a move never changes a shape's size or rotation.

import type { ReferenceOptions } from './bounds.ts'
import { areasFor, bottomOf, centreX, centreY, copy, resolveReference, rightOf, translate, union, visualBounds } from './bounds.ts'
import type { Axis, Box, Extent, Rect, SlideSize } from './types.ts'

export type AlignEdge = 'left' | 'centre' | 'right' | 'top' | 'middle' | 'bottom'

export interface AlignOptions extends ReferenceOptions {
  /** Move the shapes together, keeping their layout, so the group's bounds line up with the reference. */
  asGroup?: boolean
}

function delta(v: Rect, r: Rect, edge: AlignEdge): [number, number] {
  switch (edge) {
    case 'left': return [r.left - v.left, 0]
    case 'centre': return [centreX(r) - centreX(v), 0]
    case 'right': return [rightOf(r) - rightOf(v), 0]
    case 'top': return [0, r.top - v.top]
    case 'middle': return [0, centreY(r) - centreY(v)]
    case 'bottom': return [0, bottomOf(r) - bottomOf(v)]
  }
}

/**
 * Lines shapes up with a reference. With 'selection' they line up with the selection's outermost edge (one shape
 * on its own lines up with the margins); with 'anchor' the anchor stays put and the others line up with it.
 */
export function align(boxes: readonly Box[], edge: AlignEdge, options: AlignOptions): Box[] {
  if (!boxes.length) return []
  const { rect, fixedId } = resolveReference(boxes, options)
  if (options.asGroup) {
    const moving = union(boxes.filter((b) => b.id !== fixedId).map(visualBounds))
    if (!moving) return boxes.map(copy)
    const [dx, dy] = delta(moving, rect, edge)
    return boxes.map((b) => (b.id === fixedId ? copy(b) : translate(b, dx, dy)))
  }
  return boxes.map((b) => {
    if (b.id === fixedId) return copy(b)
    const [dx, dy] = delta(visualBounds(b), rect, edge)
    return translate(b, dx, dy)
  })
}

interface Item {
  box: Box
  index: number
  start: number
  size: number
}

function spanOf(r: Rect, axis: Axis): { start: number; size: number } {
  return axis === 'horizontal' ? { start: r.left, size: r.width } : { start: r.top, size: r.height }
}

/** The boxes in order along an axis: by where they start, then by centre, then by their order in the input. */
function along(boxes: readonly Box[], axis: Axis): Item[] {
  return boxes
    .map((box, index) => ({ box, index, ...spanOf(visualBounds(box), axis) }))
    .sort((a, b) => a.start - b.start || a.start + a.size / 2 - (b.start + b.size / 2) || a.index - b.index)
}

/** Moves each item so its bounds start at the given position on the axis; returns boxes in input order. */
function place(boxes: readonly Box[], axis: Axis, items: readonly Item[], starts: readonly number[]): Box[] {
  const out = boxes.map(copy)
  items.forEach((it, k) => {
    const d = (starts[k] ?? it.start) - it.start
    out[it.index] = axis === 'horizontal' ? translate(it.box, d, 0) : translate(it.box, 0, d)
  })
  return out
}

/** Starts for butting items together with a fixed gap, from the first item's start. */
function stacked(items: readonly Item[], gap: number): number[] {
  const starts: number[] = []
  let pos = items[0]?.start ?? 0
  for (const it of items) {
    starts.push(pos)
    pos += it.size + gap
  }
  return starts
}

export interface DistributeOptions {
  axis: Axis
  /** Equal gaps between the shapes (the default), or equal distances between their centres. */
  spacing?: 'gaps' | 'centres'
  /** Between the outer shapes of the selection (the default), or across an area of the slide. */
  within?: Extent
  /** An exact gap in points: stacks the shapes from the first one, whatever `within` says. */
  gap?: number
  slide?: SlideSize
}

/**
 * Spreads shapes along an axis. Within the selection, the outer two stay where they are and three or more shapes
 * are needed; across an area, the first starts at its near edge and the last ends at its far edge, and a single
 * shape is centred in it.
 */
export function distribute(boxes: readonly Box[], options: DistributeOptions): Box[] {
  const { axis } = options
  const items = along(boxes, axis)
  if (!items.length) return []
  if (options.gap !== undefined) return place(boxes, axis, items, stacked(items, options.gap))
  const within = options.within ?? 'selection'
  let lo: number
  let hi: number
  if (within === 'selection') {
    if (items.length < 3) return boxes.map(copy)
    lo = Math.min(...items.map((i) => i.start))
    hi = Math.max(...items.map((i) => i.start + i.size))
  } else {
    const area = spanOf(areasFor(options.slide)[within], axis)
    lo = area.start
    hi = area.start + area.size
  }
  const first = items[0]
  const last = items[items.length - 1]
  if (!first || !last) return boxes.map(copy)
  if (items.length === 1) return place(boxes, axis, items, [lo + (hi - lo - first.size) / 2])
  if ((options.spacing ?? 'gaps') === 'gaps') {
    const gap = (hi - lo - items.reduce((sum, i) => sum + i.size, 0)) / (items.length - 1)
    return place(boxes, axis, items, stacked(items, gap).map((s) => s - first.start + lo))
  }
  const c0 = lo + first.size / 2
  const c1 = hi - last.size / 2
  return place(boxes, axis, items, items.map((it, k) => c0 + ((c1 - c0) * k) / (items.length - 1) - it.size / 2))
}

/**
 * Makes the gaps between shapes equal, keeping their order and the first shape where it is. Every gap becomes the
 * average of the gaps there are, so a row without overlaps keeps its overall length; unlike distribute, the last
 * shape is not pinned, so overlapping shapes are pulled apart rather than squeezed.
 */
export function equaliseGaps(boxes: readonly Box[], axis: Axis): Box[] {
  const items = along(boxes, axis)
  if (items.length < 3) return boxes.map(copy)
  const gaps = items.slice(1).map((it, k) => {
    const prev = items[k]
    return prev ? it.start - (prev.start + prev.size) : 0
  })
  const gap = gaps.reduce((a, b) => a + b, 0) / gaps.length
  return place(boxes, axis, items, stacked(items, gap))
}

/** Butts shapes together edge to edge with a gap (none by default), in order, from the first one. */
export function stack(boxes: readonly Box[], axis: Axis, gap = 0): Box[] {
  if (boxes.length < 2) return boxes.map(copy)
  return distribute(boxes, { axis, gap })
}
