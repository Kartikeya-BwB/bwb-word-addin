/**
 * The geometry engine behind the arrange tools: pure functions from boxes to new boxes, in points. Nothing here
 * talks to PowerPoint; src/ppt/arrange.ts reads the selection, calls these and writes the result back.
 *
 * Every operation returns new boxes in the order it was given, untouched ones as unchanged copies, rounded to a
 * hundredth of a point. Moves work on rotated bounds (what people see) and never change size or rotation.
 */
export type { Axis, Box, Edge, Extent, Rect, Reference, SlideSize } from './types.ts'
export { GeometryError } from './types.ts'
export {
  DEFAULT_SLIDE, PT_PER_CM, areasFor, cmToPt, effectiveSize, gridColumns, largest, ptToCm, resolveReference, rightAngle, roundPt, union,
  visualBounds,
} from './bounds.ts'
export type { ReferenceOptions } from './bounds.ts'
export { align, distribute, equaliseGaps, stack } from './align.ts'
export type { AlignEdge, AlignOptions, DistributeOptions } from './align.ts'
export { matchSize, setFrame, snapToGrid, stretch } from './size.ts'
export type { Dimension, FrameValues, MatchSizeOptions, SizeSource, StretchOptions } from './size.ts'
export { grid, nudge, readingOrder, swap, tidyUp } from './layout.ts'
export type { GridOptions, TidyOptions } from './layout.ts'
