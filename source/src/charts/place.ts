// Label placement for points, by four simple rules applied in order:
//   1. put the label to the right of its point;
//   2. if that would cover another label, a point, an annotation or the edge of the plot, put it to the left;
//   3. failing that, above the point;
//   4. failing that, below it.
// A label with no free side is left out, and the chart says how many were. Nothing moves once placed, so the result
// depends only on the order labels are placed in, which is the order of the data (largest bubbles first).

import { textWidth } from './measure.ts'
import type { Anchor } from './svg.ts'
import { centredBaseline } from './svg.ts'

export interface Rect {
  x0: number
  y0: number
  x1: number
  y1: number
}

export type Side = 'right' | 'left' | 'above' | 'below'

export const SIDES: readonly Side[] = ['right', 'left', 'above', 'below']

export interface Placed {
  /** Where to write the text: x, baseline and anchor. */
  x: number
  y: number
  anchor: Anchor
  box: Rect
  side: Side
}

/** The box a line of text covers, from its cap height to its descenders. */
export function textRect(x: number, baseline: number, text: string, size: number, anchor: Anchor = 'start', bold = false): Rect {
  const width = textWidth(text, size, bold)
  const x0 = anchor === 'middle' ? x - width / 2 : anchor === 'end' ? x - width : x
  return { x0, y0: baseline - size * 0.72, x1: x0 + width, y1: baseline + size * 0.22 }
}

/** A square around a circle, for points and bubbles as obstacles. */
export const circleRect = (cx: number, cy: number, r: number): Rect => ({ x0: cx - r, y0: cy - r, x1: cx + r, y1: cy + r })

export const overlap = (a: Rect, b: Rect, pad = 0): boolean =>
  a.x0 < b.x1 + pad && b.x0 < a.x1 + pad && a.y0 < b.y1 + pad && b.y0 < a.y1 + pad

export class LabelPlacer {
  private readonly taken: Rect[]
  /** Labels that found no free side. */
  dropped = 0

  constructor(private readonly bounds: Rect, obstacles: Rect[] = [], private readonly pad = 0.75) {
    this.taken = [...obstacles]
  }

  /** Marks an area as used, such as a label placed by other means. */
  block(rect: Rect): void {
    this.taken.push(rect)
  }

  /** True when a box is inside the bounds and clear of everything placed or blocked. */
  fits(rect: Rect): boolean {
    const b = this.bounds
    if (rect.x0 < b.x0 || rect.x1 > b.x1 || rect.y0 < b.y0 || rect.y1 > b.y1) return false
    return !this.taken.some((t) => overlap(rect, t, this.pad))
  }

  /** Places a label beside a point of radius r, trying the sides in order; null when none is free. */
  place(cx: number, cy: number, r: number, text: string, size: number, options: { bold?: boolean; sides?: readonly Side[] } = {}): Placed | null {
    const gap = size * 0.3
    const bold = options.bold ?? false
    for (const side of options.sides ?? SIDES) {
      const spot = spotFor(side, cx, cy, r + gap, size)
      const box = textRect(spot.x, spot.y, text, size, spot.anchor, bold)
      if (this.fits(box)) {
        this.taken.push(box)
        return { ...spot, box, side }
      }
    }
    this.dropped += 1
    return null
  }
}

function spotFor(side: Side, cx: number, cy: number, reach: number, size: number): { x: number; y: number; anchor: Anchor } {
  switch (side) {
    case 'right':
      return { x: cx + reach, y: centredBaseline(cy, size), anchor: 'start' }
    case 'left':
      return { x: cx - reach, y: centredBaseline(cy, size), anchor: 'end' }
    case 'above':
      return { x: cx, y: cy - reach - size * 0.22, anchor: 'middle' }
    case 'below':
      return { x: cx, y: cy + reach + size * 0.72, anchor: 'middle' }
  }
}
