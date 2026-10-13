// The arrange tools in PowerPoint: read the selected shapes, run a geometry operation, write the changes back in one
// sync. Also the order of shapes (z-order) and grouping, which PowerPoint does itself.

import type {
  AlignEdge, Axis, Box, Dimension, DistributeOptions, Edge, FrameValues, GridOptions, Reference, SizeSource, SlideSize,
} from '../geometry/index.ts'
import {
  DEFAULT_SLIDE, GeometryError, align, distribute, equaliseGaps, grid, matchSize, nudge, setFrame, snapToGrid, stack, stretch, swap, tidyUp,
} from '../geometry/index.ts'
import { UserFacingError } from '../word/run.ts'
import { requirePpt, runPpt, supportsPpt } from './run.ts'
import { anchorHint } from './selection.ts'

export interface ArrangeInput {
  /** The selected shapes: those whose order of selection is known first, in that order, then the rest. */
  boxes: Box[]
  /** The anchor: set by hand (maybe outside the selection) or the first selected. Null leaves it to the largest. */
  anchor: Box | null
  slide: SlideSize
}

interface Selected {
  shape: PowerPoint.Shape
  box: Box
}

const FIELDS = ['left', 'top', 'width', 'height'] as const

const plural = (n: number, word: string): string => `${n} ${word}${n === 1 ? '' : 's'}`

function boxOf(shape: PowerPoint.Shape, rotation: boolean): Box {
  return { id: shape.id, left: shape.left, top: shape.top, width: shape.width, height: shape.height, rotation: rotation ? shape.rotation : 0 }
}

export interface ArrangeOptions {
  /** The fewest shapes the operation needs, counting the anchor when it is used and sits outside the selection. */
  min: number
  /** What to tell people who selected fewer. */
  need: string
  /** The operation lines up with or sizes from the anchor, which may then be outside the selection. */
  anchor?: boolean
}

/** Runs an operation on the selection and writes back what changed, in one sync. */
export async function arrange(op: (input: ArrangeInput) => Box[], options: ArrangeOptions): Promise<{ changed: number; total: number }> {
  const { min, need } = options
  requirePpt('1.5', 'Arranging shapes')
  const rotation = supportsPpt('1.10')
  return runPpt(async (context) => {
    const presentation = context.presentation
    const shapes = presentation.getSelectedShapes()
    shapes.load(`items/id,items/left,items/top,items/width,items/height${rotation ? ',items/rotation' : ''}`)
    const slides = presentation.getSelectedSlides()
    slides.load('items/id')
    const page = rotation ? presentation.pageSetup : null
    page?.load('slideWidth,slideHeight')
    await context.sync()

    const slide = slides.items[0] ?? null
    const selected: Selected[] = shapes.items.map((shape) => ({ shape, box: boxOf(shape, rotation) }))
    const hint = anchorHint(slide?.id ?? null, selected.map((s) => s.box.id))
    let anchor = selected.find((s) => s.box.id === hint.anchorId)?.box ?? null
    if (!anchor && options.anchor && hint.explicit && hint.anchorId && slide) {
      const outside = slide.shapes.getItemOrNullObject(hint.anchorId)
      outside.load(`id,left,top,width,height${rotation ? ',rotation' : ''}`)
      await context.sync()
      if (!outside.isNullObject) anchor = boxOf(outside, rotation)
    }
    const extra = anchor && !selected.some((s) => s.box.id === anchor?.id) ? 1 : 0
    if (selected.length + extra < min || !selected.length) throw new UserFacingError(need)

    const rank = new Map(hint.order.map((id, i) => [id, i]))
    const ordered = [...selected].sort((a, b) => (rank.get(a.box.id) ?? Number.MAX_SAFE_INTEGER) - (rank.get(b.box.id) ?? Number.MAX_SAFE_INTEGER))
    const slideSize = page ? { width: page.slideWidth, height: page.slideHeight } : DEFAULT_SLIDE
    let result: Box[]
    try {
      result = op({ boxes: ordered.map((s) => s.box), anchor, slide: slideSize })
    } catch (error) {
      if (error instanceof GeometryError) throw new UserFacingError(error.message)
      throw error
    }

    const byId = new Map(selected.map((s) => [s.box.id, s]))
    let changed = 0
    for (const next of result) {
      const s = byId.get(next.id)
      if (!s) continue
      const fields = FIELDS.filter((k) => Math.abs(next[k] - s.box[k]) >= 0.005)
      for (const k of fields) s.shape[k] = next[k]
      if (fields.length) changed += 1
    }
    if (changed) await context.sync()
    return { changed, total: selected.length }
  })
}

function report(result: { changed: number }, done: string): string {
  return result.changed ? `${done} ${plural(result.changed, 'shape')}.` : 'Nothing to change: the shapes are already in place.'
}

const SELECT_SOME = 'Select the shapes first.'

/** Lines shapes up with a reference; asGroup moves them together. One shape against the selection uses the margins. */
export async function alignShapes(edge: AlignEdge, reference: Reference, asGroup = false): Promise<string> {
  const r = await arrange((i) => align(i.boxes, edge, { reference, anchor: i.anchor, slide: i.slide, asGroup }), {
    min: 1, need: SELECT_SOME, anchor: reference === 'anchor',
  })
  return report(r, asGroup ? 'Moved' : 'Lined up')
}

export async function distributeShapes(options: Omit<DistributeOptions, 'slide'>): Promise<string> {
  const between = (options.within ?? 'selection') === 'selection' && options.gap === undefined
  const r = await arrange((i) => distribute(i.boxes, { ...options, slide: i.slide }), {
    min: between ? 3 : 1, need: between ? 'Select three or more shapes to space them out between the outer two.' : SELECT_SOME,
  })
  return report(r, 'Spaced out')
}

export async function equaliseShapeGaps(axis: Axis): Promise<string> {
  const r = await arrange((i) => equaliseGaps(i.boxes, axis), { min: 3, need: 'Select three or more shapes to even out their gaps.' })
  return report(r, 'Evened out the gaps between')
}

export async function matchShapeSize(dimension: Dimension, source: SizeSource, keepAspect = false): Promise<string> {
  const r = await arrange((i) => matchSize(i.boxes, { dimension, source, anchor: i.anchor, keepAspect }), {
    min: 2, need: 'Select two or more shapes. The anchor, or the first one you select, sets the size.', anchor: source === 'anchor',
  })
  return report(r, 'Resized')
}

/** Sets the frame of every selected shape exactly, in points (the pane converts from centimetres). */
export async function setShapeFrame(values: FrameValues): Promise<string> {
  const r = await arrange((i) => setFrame(i.boxes, values), { min: 1, need: SELECT_SOME })
  return report(r, 'Placed')
}

export async function stackShapes(axis: Axis, gap = 0): Promise<string> {
  const r = await arrange((i) => stack(i.boxes, axis, gap), { min: 2, need: 'Select two or more shapes to stack.' })
  return report(r, 'Stacked')
}

export async function gridShapes(options: GridOptions): Promise<string> {
  const r = await arrange((i) => grid(i.boxes, options), { min: 2, need: 'Select two or more shapes to lay out in a grid.' })
  return report(r, 'Arranged')
}

export async function tidyShapes(): Promise<string> {
  const r = await arrange((i) => tidyUp(i.boxes), { min: 2, need: 'Select two or more shapes to tidy up.' })
  return report(r, 'Tidied')
}

export async function swapShapes(): Promise<string> {
  const r = await arrange((i) => swap(i.boxes), { min: 2, need: 'Select two shapes to swap.' })
  return r.changed ? 'Swapped the two shapes.' : 'Nothing to change: the shapes are in the same place.'
}

export async function stretchShapes(edge: Edge, reference: Reference): Promise<string> {
  const r = await arrange((i) => stretch(i.boxes, edge, { reference, anchor: i.anchor, slide: i.slide }), {
    min: 1, need: SELECT_SOME, anchor: reference === 'anchor',
  })
  return report(r, 'Stretched')
}

export async function snapShapes(): Promise<string> {
  const r = await arrange((i) => snapToGrid(i.boxes, { slide: i.slide }), { min: 1, need: SELECT_SOME })
  return report(r, 'Snapped')
}

export async function nudgeShapes(dx: number, dy: number): Promise<string> {
  const r = await arrange((i) => nudge(i.boxes, dx, dy), { min: 1, need: SELECT_SOME })
  return report(r, 'Moved')
}

/** The frame of the anchor, or of the first selected shape, to show in the pane. Null when nothing is selected. */
export async function readShapeFrame(): Promise<Required<FrameValues> | null> {
  let frame: Required<FrameValues> | null = null
  await arrange((i) => {
    const first = i.boxes.find((b) => b.id === i.anchor?.id) ?? i.boxes[0]
    if (first) frame = { left: first.left, top: first.top, width: first.width, height: first.height }
    return i.boxes
  }, { min: 0, need: SELECT_SOME }).catch((error: unknown) => {
    if (!(error instanceof UserFacingError) || error.message !== SELECT_SOME) throw error
  })
  return frame
}

export type ZOrder = 'BringToFront' | 'SendToBack' | 'BringForward' | 'SendBackward'

const Z_DONE: Record<ZOrder, string> = {
  BringToFront: 'Brought to the front', SendToBack: 'Sent to the back', BringForward: 'Brought forward', SendBackward: 'Sent backward',
}

/** Changes the order of the selected shapes, keeping their order among themselves. */
export async function orderShapes(position: ZOrder): Promise<string> {
  requirePpt('1.8', 'Changing the order of shapes')
  return runPpt(async (context) => {
    const shapes = context.presentation.getSelectedShapes()
    shapes.load('items/id,items/zOrderPosition')
    await context.sync()
    if (!shapes.items.length) throw new UserFacingError(SELECT_SOME)
    // To the front, the back-most goes first; one step forward, the front-most goes first. And the reverse for back.
    const backFirst = position === 'BringToFront' || position === 'SendBackward'
    const ordered = [...shapes.items].sort((a, b) => (backFirst ? 1 : -1) * (a.zOrderPosition - b.zOrderPosition))
    ordered.forEach((s) => s.setZOrder(position))
    await context.sync()
    return `${Z_DONE[position]}: ${plural(ordered.length, 'shape')}.`
  })
}

export async function groupShapes(): Promise<string> {
  requirePpt('1.8', 'Grouping shapes')
  return runPpt(async (context) => {
    const shapes = context.presentation.getSelectedShapes()
    shapes.load('items/id')
    await context.sync()
    if (shapes.items.length < 2) throw new UserFacingError('Select two or more shapes to group.')
    shapes.group()
    await context.sync()
    return `Grouped ${plural(shapes.items.length, 'shape')}.`
  })
}

export async function ungroupShapes(): Promise<string> {
  requirePpt('1.8', 'Ungrouping shapes')
  return runPpt(async (context) => {
    const shapes = context.presentation.getSelectedShapes()
    shapes.load('items/id,items/type')
    await context.sync()
    const groups = shapes.items.filter((s) => s.type === 'Group')
    if (!groups.length) throw new UserFacingError('Select a group to ungroup.')
    groups.forEach((g) => g.group.ungroup())
    await context.sync()
    return `Ungrouped ${plural(groups.length, 'group')}.`
  })
}
