// Turns the element layouts of specs.ts into PowerPoint shapes on the current slide: formatted, named, tagged with
// BWB_ELEMENT, grouped when PowerPoint can, and selected.

import { TAGS } from '../brand/slideGrid.ts'
import { UserFacingError } from '../word/run.ts'
import { requirePpt, runPpt, supportsPpt } from './run.ts'
import type { ElementSpec, FontSpec, ShapeSpec, TableSpec, TextSpec } from './specs.ts'

/** The slide in the editing pane (the first of the selected slides). */
export async function currentSlide(context: PowerPoint.RequestContext): Promise<PowerPoint.Slide> {
  const slides = context.presentation.getSelectedSlides()
  slides.load('items/id')
  await context.sync()
  const slide = slides.items[0]
  if (!slide) throw new UserFacingError('Click a slide first, then try again.')
  return slide
}

function setFont(font: PowerPoint.ShapeFont, spec: FontSpec): void {
  font.name = spec.name
  font.size = spec.size
  font.color = `#${spec.colour}`
  font.bold = spec.bold ?? false
  font.italic = spec.italic ?? false
}

const joined = (spec: TextSpec): string => spec.paragraphs.map((p) => p.text).join('\n')

/**
 * Sets up the frame first (autofit off, so a new text box keeps its size), then writes the text, then each
 * paragraph's font together with its paragraph mark.
 */
export function applyText(frame: PowerPoint.TextFrame, spec: TextSpec): void {
  frame.autoSizeSetting = 'AutoSizeNone'
  frame.wordWrap = spec.wrap
  frame.verticalAlignment = spec.vertical
  frame.leftMargin = spec.margins.left
  frame.rightMargin = spec.margins.right
  frame.topMargin = spec.margins.top
  frame.bottomMargin = spec.margins.bottom
  const range = frame.textRange
  range.text = joined(spec)
  range.paragraphFormat.horizontalAlignment = spec.align
  let at = 0
  spec.paragraphs.forEach((p, i) => {
    const length = p.text.length + (i < spec.paragraphs.length - 1 ? 1 : 0)
    if (length) setFont(range.getSubstring(at, length).font, p.font)
    at += p.text.length + 1
  })
}

/** Adds one shape from its layout. Rotation needs PowerPointApi 1.10 and is skipped without it. */
export function addShape(shapes: PowerPoint.ShapeCollection, spec: ShapeSpec, canRotate: boolean): PowerPoint.Shape {
  const at = { left: spec.left, top: spec.top, width: spec.width, height: spec.height }
  const shape = spec.kind === 'text'
    ? shapes.addTextBox(spec.text ? joined(spec.text) : '', at)
    : spec.kind === 'line' ? shapes.addLine('Straight', at) : shapes.addGeometricShape(spec.geometry, at)
  shape.name = spec.name
  if (spec.rotation && canRotate) shape.rotation = spec.rotation
  if (spec.kind !== 'line') {
    if (spec.fill) shape.fill.setSolidColor(`#${spec.fill}`)
    else shape.fill.clear()
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

export interface Inserted {
  ids: string[]
  groupId: string | null
}

/** Adds an element to the current slide, tags every shape (and the group) and selects the result. */
export async function insertElement(spec: ElementSpec, tags: Record<string, string> = {}): Promise<Inserted> {
  requirePpt('1.5', 'Inserting BwB elements')
  const canRotate = supportsPpt('1.10')
  const canGroup = spec.group && spec.shapes.length > 1 && supportsPpt('1.8')
  return runPpt(async (context) => {
    const slide = await currentSlide(context)
    const tag = (shape: PowerPoint.Shape): void => {
      shape.tags.add(TAGS.element, spec.kind)
      for (const [key, value] of Object.entries(tags)) shape.tags.add(key, value)
      shape.load('id')
    }
    const shapes = spec.shapes.map((s) => {
      const shape = addShape(slide.shapes, s, canRotate)
      tag(shape)
      return shape
    })
    await context.sync()
    const ids = shapes.map((s) => s.id)
    let groupId: string | null = null
    if (canGroup) {
      const g = slide.shapes.addGroup(ids)
      g.name = spec.name
      tag(g)
      await context.sync()
      groupId = g.id
    }
    slide.setSelectedShapes(groupId ? [groupId] : ids)
    await context.sync()
    return { ids, groupId }
  })
}

/** Adds the BwB table (PowerPointApi 1.8; with 1.9 it also drops PowerPoint's own table style) and selects it. */
export async function insertTable(spec: TableSpec): Promise<Inserted> {
  requirePpt('1.8', 'The BwB table')
  const plain = supportsPpt('1.9')
  return runPpt(async (context) => {
    const slide = await currentSlide(context)
    const rule: PowerPoint.BorderProperties = { color: `#${spec.rule.colour}`, weight: spec.rule.weight, dashStyle: 'Solid' }
    const cell = (part: TableSpec['header']): PowerPoint.TableCellProperties => ({
      fill: { color: `#${part.fill}` },
      font: { name: part.font.name, size: part.font.size, color: `#${part.font.colour}`, bold: part.font.bold ?? false },
      borders: { top: rule, bottom: rule, left: rule, right: rule },
    })
    const options: PowerPoint.TableAddOptions = {
      values: spec.values,
      left: spec.left,
      top: spec.top,
      width: spec.width,
      specificCellProperties: spec.values.map((row, r) => row.map(() => cell(r === 0 ? spec.header : spec.body))),
      ...(plain ? { style: 'NoStyleNoGrid' as const } : {}),
    }
    const table = slide.shapes.addTable(spec.rows, spec.columns, options)
    table.name = 'BwB table'
    table.tags.add(TAGS.element, 'table')
    table.load('id')
    await context.sync()
    slide.setSelectedShapes([table.id])
    await context.sync()
    return { ids: [table.id], groupId: null }
  })
}

export interface TaggedShape {
  slide: PowerPoint.Slide
  shape: PowerPoint.Shape
  value: string
}

/** The top-level shapes on these slides that carry a tag, with its value. Two syncs, whatever the number of slides. */
export async function taggedShapes(context: PowerPoint.RequestContext, slides: readonly PowerPoint.Slide[], key: string): Promise<TaggedShape[]> {
  const lists = slides.map((slide) => {
    const shapes = slide.shapes
    shapes.load('items/id')
    return { slide, shapes }
  })
  await context.sync()
  const probes = lists.flatMap(({ slide, shapes }) =>
    shapes.items.map((shape) => {
      const tag = shape.tags.getItemOrNullObject(key)
      tag.load('value')
      return { slide, shape, tag }
    }),
  )
  await context.sync()
  return probes.filter((p) => !p.tag.isNullObject).map((p) => ({ slide: p.slide, shape: p.shape, value: p.tag.value }))
}
