// Rotation-aware bounds, the slide's reference areas and the resizing primitives the operations build on.

import { COLUMNS, CONTENT_AREA, GUTTER, MARGINS, SLIDE } from '../brand/slideGrid.ts'
import type { Box, Rect, Reference, SlideSize } from './types.ts'

export const DEFAULT_SLIDE: SlideSize = { width: SLIDE.width, height: SLIDE.height }

export const PT_PER_CM = 72 / 2.54
export const cmToPt = (cm: number): number => cm * PT_PER_CM
export const ptToCm = (pt: number): number => pt / PT_PER_CM

/** Rounds to a hundredth of a point, finer than PowerPoint shows, without negative zero. */
export function roundPt(x: number): number {
  const r = Math.round(x * 100) / 100
  return r === 0 ? 0 : r
}

export const rightOf = (r: Rect): number => r.left + r.width
export const bottomOf = (r: Rect): number => r.top + r.height
export const centreX = (r: Rect): number => r.left + r.width / 2
export const centreY = (r: Rect): number => r.top + r.height / 2

export function normaliseRotation(rotation = 0): number {
  const r = ((rotation % 360) + 360) % 360
  return Math.abs(r - 360) < 1e-9 ? 0 : r
}

/** 0, 90, 180 or 270 when the rotation is a right angle (to a hundredth of a degree), otherwise null. */
export function rightAngle(rotation = 0): 0 | 90 | 180 | 270 | null {
  const r = normaliseRotation(rotation)
  for (const a of [0, 90, 180, 270, 360] as const) if (Math.abs(r - a) < 0.01) return a === 360 ? 0 : a
  return null
}

/** |cos| and |sin| of the rotation, exact for right angles. */
function turn(rotation = 0): { c: number; s: number } {
  const right = rightAngle(rotation)
  if (right === 0 || right === 180) return { c: 1, s: 0 }
  if (right === 90 || right === 270) return { c: 0, s: 1 }
  const t = (normaliseRotation(rotation) * Math.PI) / 180
  return { c: Math.abs(Math.cos(t)), s: Math.abs(Math.sin(t)) }
}

/** The box a shape takes up on the slide once rotated: what people see, and what alignment works with. */
export function visualBounds(b: Box): Rect {
  const { c, s } = turn(b.rotation)
  const width = b.width * c + b.height * s
  const height = b.width * s + b.height * c
  return { left: b.left + (b.width - width) / 2, top: b.top + (b.height - height) / 2, width, height }
}

export function union(rects: readonly Rect[]): Rect | null {
  if (!rects.length) return null
  const left = Math.min(...rects.map((r) => r.left))
  const top = Math.min(...rects.map((r) => r.top))
  return { left, top, width: Math.max(...rects.map(rightOf)) - left, height: Math.max(...rects.map(bottomOf)) - top }
}

export const copy = (b: Box): Box => ({ ...b })

/** Moves a box. A move of nothing returns the box unchanged, so untouched shapes are never rewritten. */
export function translate(b: Box, dx: number, dy: number): Box {
  if (Math.abs(dx) < 1e-9 && Math.abs(dy) < 1e-9) return copy(b)
  return { ...b, left: roundPt(b.left + dx), top: roundPt(b.top + dy) }
}

/** Moves a box so the top-left corner of its bounds is at (left, top); a missing coordinate stays. */
export function moveBoundsTo(b: Box, left?: number, top?: number): Box {
  const v = visualBounds(b)
  return translate(b, left === undefined ? 0 : left - v.left, top === undefined ? 0 : top - v.top)
}

function rounded(b: Box): Box {
  return { ...b, left: roundPt(b.left), top: roundPt(b.top), width: roundPt(b.width), height: roundPt(b.height) }
}

/** Width and height as the shape reads on the slide: swapped for a quarter turn, otherwise the frame's own. */
export function effectiveSize(b: Box): { width: number; height: number } {
  const r = rightAngle(b.rotation)
  return r === 90 || r === 270 ? { width: b.height, height: b.width } : { width: b.width, height: b.height }
}

/** Gives a box a new effective size, keeping the top-left corner of its bounds where it was. */
export function resizeEffective(b: Box, width: number, height: number): Box {
  const r = rightAngle(b.rotation)
  const frame = r === 90 || r === 270 ? { width: height, height: width } : { width, height }
  const before = visualBounds(b)
  const sized = { ...b, ...frame }
  const after = visualBounds(sized)
  return rounded({ ...sized, left: sized.left + before.left - after.left, top: sized.top + before.top - after.top })
}

/**
 * Resizes a box so its bounds are `width` wide, keeping one side of the bounds in place and the vertical centre.
 * A rotated frame changes along whichever of its sides lies closer to horizontal. Null when no frame fits.
 */
export function setBoundsWidth(b: Box, width: number, keep: 'left' | 'right'): Box | null {
  const { c, s } = turn(b.rotation)
  const frame = c >= s ? { width: (width - b.height * s) / c, height: b.height } : { width: b.width, height: (width - b.width * c) / s }
  if (!(frame.width > 0) || !(frame.height > 0) || !Number.isFinite(frame.width + frame.height)) return null
  const before = visualBounds(b)
  const sized = { ...b, ...frame }
  const after = visualBounds(sized)
  const dx = keep === 'left' ? before.left - after.left : rightOf(before) - rightOf(after)
  return rounded({ ...sized, left: sized.left + dx, top: sized.top + centreY(before) - centreY(after) })
}

/** The vertical twin of setBoundsWidth: bounds `height` tall, one side kept, the horizontal centre kept. */
export function setBoundsHeight(b: Box, height: number, keep: 'top' | 'bottom'): Box | null {
  const { c, s } = turn(b.rotation)
  const frame = c >= s ? { width: b.width, height: (height - b.width * s) / c } : { width: (height - b.height * c) / s, height: b.height }
  if (!(frame.width > 0) || !(frame.height > 0) || !Number.isFinite(frame.width + frame.height)) return null
  const before = visualBounds(b)
  const sized = { ...b, ...frame }
  const after = visualBounds(sized)
  const dy = keep === 'top' ? before.top - after.top : bottomOf(before) - bottomOf(after)
  return rounded({ ...sized, left: sized.left + centreX(before) - centreX(after), top: sized.top + dy })
}

/** The slide, the area inside the BwB margins and the content area, for a slide of any size. */
export function areasFor(slide: SlideSize = DEFAULT_SLIDE): { slide: Rect; margins: Rect; content: Rect } {
  const below = SLIDE.height - bottomOf(CONTENT_AREA)
  const beside = SLIDE.width - rightOf(CONTENT_AREA)
  return {
    slide: { left: 0, top: 0, width: slide.width, height: slide.height },
    margins: { left: MARGINS.left, top: MARGINS.top, width: slide.width - MARGINS.left - MARGINS.right, height: slide.height - MARGINS.top - MARGINS.bottom },
    content: { left: CONTENT_AREA.left, top: CONTENT_AREA.top, width: slide.width - CONTENT_AREA.left - beside, height: slide.height - CONTENT_AREA.top - below },
  }
}

/** The 12 BwB columns inside the margins of a slide of any size: the same edges as columnEdges() on 960 by 540. */
export function gridColumns(slide: SlideSize = DEFAULT_SLIDE): Array<{ left: number; right: number }> {
  const m = areasFor(slide).margins
  const width = (m.width - (COLUMNS - 1) * GUTTER) / COLUMNS
  return Array.from({ length: COLUMNS }, (_, i) => {
    const left = m.left + i * (width + GUTTER)
    return { left, right: left + width }
  })
}

/** The largest box by the area of its bounds; the first of equals. */
export function largest(boxes: readonly Box[]): Box | undefined {
  let best: Box | undefined
  let bestArea = -1
  for (const b of boxes) {
    const v = visualBounds(b)
    const area = v.width * v.height
    if (area > bestArea + 1e-9) {
      best = b
      bestArea = area
    }
  }
  return best
}

export interface ReferenceOptions {
  reference: Reference
  /** The key object. It may be outside `boxes`. Without one, the largest box is the anchor. */
  anchor?: Box | null
  slide?: SlideSize
}

/** The rectangle to line up with, and the id of the shape that stays put (the anchor), if any. */
export function resolveReference(boxes: readonly Box[], options: ReferenceOptions): { rect: Rect; fixedId: string | null } {
  const areas = areasFor(options.slide)
  switch (options.reference) {
    case 'selection':
      // One shape has nothing to line up with, so it lines up with the margins.
      return { rect: boxes.length < 2 ? areas.margins : (union(boxes.map(visualBounds)) ?? areas.margins), fixedId: null }
    case 'anchor': {
      const anchor = options.anchor ?? largest(boxes)
      return anchor ? { rect: visualBounds(anchor), fixedId: anchor.id } : { rect: areas.margins, fixedId: null }
    }
    case 'slide':
      return { rect: areas.slide, fixedId: null }
    case 'margins':
      return { rect: areas.margins, fixedId: null }
    case 'content':
      return { rect: areas.content, fixedId: null }
  }
}
