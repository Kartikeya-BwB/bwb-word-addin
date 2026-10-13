// Drawings made of native shapes (diagrams, org charts): plain layouts in points, turned into PowerPoint shapes on
// the current slide, tagged BWB_ELEMENT, grouped when PowerPoint can (1.8) and selected.

import { TAGS } from '../../brand/slideGrid.ts'
import { applyText, currentSlide } from '../render.ts'
import { requirePpt, runPpt, supportsPpt } from '../run.ts'
import type { TextSpec } from '../specs.ts'

export interface DrawShape {
  kind: 'geometric' | 'text' | 'line'
  /** A PowerPoint geometric shape type, such as Rectangle, RoundRectangle, Triangle or Ellipse. */
  geometry?: `${PowerPoint.GeometricShapeType}`
  name: string
  left: number
  top: number
  width: number
  height: number
  /** Degrees clockwise. Needs PowerPointApi 1.10. */
  rotation?: number
  /** Leave the shape out when PowerPoint cannot turn it, as for arrows that only make sense turned. */
  requiresRotation?: boolean
  /** Hex without '#', and transparency from 0 to 1; null for no fill. */
  fill?: { colour: string; transparency?: number } | null
  line?: { colour: string; weight: number } | null
  text?: TextSpec
}

export interface Drawn {
  ids: string[]
  groupId: string | null
  /** Shapes left out because they needed rotation. */
  skipped: number
}

/** A straight line between two points. Lines that rise to the right are drawn turned a quarter, which needs 1.10. */
export function segment(name: string, x1: number, y1: number, x2: number, y2: number, line: { colour: string; weight: number }): DrawShape {
  // Office.js draws a line from the top-left to the bottom-right of its box and refuses negative sizes, so a line
  // that rises to the right is a falling line in a box turned 90 degrees about its centre.
  const [ax, ay, bx, by] = x1 <= x2 ? [x1, y1, x2, y2] : [x2, y2, x1, y1]
  const dx = bx - ax
  const dy = by - ay
  if (dy >= 0 || Math.abs(dx) < 0.01) {
    const top = Math.min(ay, by)
    return { kind: 'line', name, left: ax, top, width: dx, height: Math.abs(dy), line }
  }
  const mx = (ax + bx) / 2
  const my = (ay + by) / 2
  return { kind: 'line', name, left: mx + dy / 2, top: my - dx / 2, width: -dy, height: dx, rotation: 90, requiresRotation: true, line }
}

function draw(shapes: PowerPoint.ShapeCollection, spec: DrawShape, canRotate: boolean): PowerPoint.Shape {
  const at = { left: spec.left, top: spec.top, width: Math.max(0, spec.width), height: Math.max(0, spec.height) }
  const shape = spec.kind === 'line'
    ? shapes.addLine('Straight', at)
    : spec.kind === 'text' ? shapes.addTextBox(spec.text?.paragraphs.map((p) => p.text).join('\n') ?? '', at) : shapes.addGeometricShape(spec.geometry ?? 'Rectangle', at)
  shape.name = spec.name
  if (spec.rotation && canRotate) shape.rotation = spec.rotation
  if (spec.kind !== 'line') {
    if (spec.fill) {
      shape.fill.setSolidColor(`#${spec.fill.colour}`)
      if (spec.fill.transparency) shape.fill.transparency = spec.fill.transparency
    } else shape.fill.clear()
  }
  if (spec.line) {
    shape.lineFormat.visible = true
    shape.lineFormat.color = `#${spec.line.colour}`
    shape.lineFormat.weight = spec.line.weight
    shape.lineFormat.dashStyle = 'Solid'
  } else {
    shape.lineFormat.visible = false
  }
  if (spec.text) applyText(shape.textFrame, spec.text)
  return shape
}

/** Draws the shapes on the current slide, back to front, tags each one and the group, and selects the result. */
export async function insertDrawing(kind: string, name: string, specs: readonly DrawShape[], feature: string): Promise<Drawn> {
  requirePpt('1.5', feature)
  const canRotate = supportsPpt('1.10')
  const canGroup = supportsPpt('1.8')
  const usable = specs.filter((s) => canRotate || !s.requiresRotation)
  return runPpt(async (context) => {
    const slide = await currentSlide(context)
    const tag = (shape: PowerPoint.Shape): void => {
      shape.tags.add(TAGS.element, kind)
      shape.load('id')
    }
    const shapes = usable.map((s) => {
      const shape = draw(slide.shapes, s, canRotate)
      tag(shape)
      return shape
    })
    await context.sync()
    const ids = shapes.map((s) => s.id)
    let groupId: string | null = null
    if (canGroup && ids.length > 1) {
      const g = slide.shapes.addGroup(ids)
      g.name = name
      tag(g)
      await context.sync()
      groupId = g.id
    }
    slide.setSelectedShapes(groupId ? [groupId] : ids)
    await context.sync()
    return { ids, groupId, skipped: specs.length - usable.length }
  })
}
