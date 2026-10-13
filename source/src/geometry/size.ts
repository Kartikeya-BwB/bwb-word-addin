// Sizing shapes: same size, exact frames, stretching an edge to a reference and snapping to the BwB grid.

import type { ReferenceOptions } from './bounds.ts'
import {
  bottomOf, copy, effectiveSize, gridColumns, largest, resizeEffective, resolveReference, rightOf, roundPt, setBoundsHeight,
  setBoundsWidth, translate, visualBounds,
} from './bounds.ts'
import type { Box, Edge, SlideSize } from './types.ts'
import { GeometryError } from './types.ts'

export type Dimension = 'width' | 'height' | 'both'

/** Where the size comes from: the anchor, the largest or smallest shape, or the first in the given order. */
export type SizeSource = 'anchor' | 'largest' | 'smallest' | 'first'

export interface MatchSizeOptions {
  dimension: Dimension
  source: SizeSource
  /** For source 'anchor'. It may be outside `boxes`; without one, the largest box is the anchor. */
  anchor?: Box | null
  /** When matching one dimension, scale the other to keep each shape's proportions. */
  keepAspect?: boolean
}

const area = (s: { width: number; height: number }): number => s.width * s.height

/**
 * Gives shapes the same width, height or both. Sizes are as the shapes read on the slide (a quarter-turned shape's
 * width is its frame's height). 'largest' and 'smallest' take the extreme width or height on their own; for both
 * dimensions they take the size of the largest or smallest shape. Each shape keeps the top-left of its bounds.
 */
export function matchSize(boxes: readonly Box[], options: MatchSizeOptions): Box[] {
  if (!boxes.length) return []
  const sizes = boxes.map(effectiveSize)
  let target: { width: number; height: number } | undefined
  let sourceId: string | null = null
  if (options.source === 'anchor' || options.source === 'first') {
    const from = options.source === 'anchor' ? (options.anchor ?? largest(boxes)) : boxes[0]
    if (from) {
      target = effectiveSize(from)
      sourceId = from.id
    }
  } else {
    const pick = options.source === 'largest' ? Math.max : Math.min
    if (options.dimension === 'both') {
      const best = pick(...sizes.map(area))
      const index = sizes.findIndex((s) => area(s) === best)
      target = sizes[index]
      sourceId = boxes[index]?.id ?? null
    } else {
      target = { width: pick(...sizes.map((s) => s.width)), height: pick(...sizes.map((s) => s.height)) }
    }
  }
  if (!target) return boxes.map(copy)
  const { width: tw, height: th } = target
  return boxes.map((b, i) => {
    const s = sizes[i]
    if (b.id === sourceId || !s) return copy(b)
    let w = options.dimension === 'height' ? s.width : tw
    let h = options.dimension === 'width' ? s.height : th
    if (options.keepAspect && options.dimension === 'width' && s.width > 0) h = (s.height * tw) / s.width
    if (options.keepAspect && options.dimension === 'height' && s.height > 0) w = (s.width * th) / s.height
    if (Math.abs(w - s.width) < 1e-9 && Math.abs(h - s.height) < 1e-9) return copy(b)
    return resizeEffective(b, w, h)
  })
}

export interface FrameValues {
  left?: number
  top?: number
  width?: number
  height?: number
}

/** Sets the frame exactly, as PowerPoint's Size and Position pane does: the unrotated frame, in points. */
export function setFrame(boxes: readonly Box[], values: FrameValues): Box[] {
  if ((values.width !== undefined && !(values.width > 0)) || (values.height !== undefined && !(values.height > 0))) {
    throw new GeometryError('Width and height must be more than zero.')
  }
  return boxes.map((b) => {
    const next = { ...b }
    for (const key of ['left', 'top', 'width', 'height'] as const) {
      const v = values[key]
      if (v !== undefined && Math.abs(v - b[key]) >= 0.005) next[key] = roundPt(v)
    }
    return next
  })
}

export interface StretchOptions extends ReferenceOptions {}

/**
 * Moves one edge of each shape to the same edge of a reference, keeping the opposite edge: right to the anchor's
 * right edge, top to the content area's top and so on. The anchor stays as it is. A shape whose edge would pass its
 * opposite edge is left alone.
 */
export function stretch(boxes: readonly Box[], edge: Edge, options: StretchOptions): Box[] {
  const { rect, fixedId } = resolveReference(boxes, options)
  return boxes.map((b) => {
    if (b.id === fixedId) return copy(b)
    const v = visualBounds(b)
    let next: Box | null
    switch (edge) {
      case 'right': next = setBoundsWidth(b, rightOf(rect) - v.left, 'left'); break
      case 'left': next = setBoundsWidth(b, rightOf(v) - rect.left, 'right'); break
      case 'bottom': next = setBoundsHeight(b, bottomOf(rect) - v.top, 'top'); break
      case 'top': next = setBoundsHeight(b, bottomOf(v) - rect.top, 'bottom'); break
    }
    return next ?? copy(b)
  })
}

function nearest(values: readonly number[], x: number): number {
  let best = 0
  values.forEach((v, i) => {
    if (Math.abs(v - x) < Math.abs((values[best] ?? Infinity) - x) - 1e-9) best = i
  })
  return best
}

/**
 * Snaps the left and right edges of each shape to the nearest lines of the BwB 12-column grid: the left edge to the
 * start of a column, the right edge to the end of one, so shapes span whole columns and gutters fall between them.
 * A shape narrower than a column takes one column.
 */
export function snapToGrid(boxes: readonly Box[], options: { slide?: SlideSize } = {}): Box[] {
  const columns = gridColumns(options.slide)
  const starts = columns.map((c) => c.left)
  const ends = columns.map((c) => c.right)
  return boxes.map((b) => {
    const v = visualBounds(b)
    const first = nearest(starts, v.left)
    const last = Math.max(first, nearest(ends, rightOf(v)))
    const left = starts[first] ?? v.left
    const right = ends[last] ?? rightOf(v)
    const sized = Math.abs(right - left - v.width) < 0.005 ? b : (setBoundsWidth(b, right - left, 'left') ?? b)
    return translate(sized, left - visualBounds(sized).left, 0)
  })
}
