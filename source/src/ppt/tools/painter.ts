// The format painter: pick up the format of one shape, then apply the parts you choose to shapes on any slide.
// Position and size need PowerPointApi 1.4; rotation and shape adjustments need 1.10.

import { UserFacingError } from '../../word/run.ts'
import { requirePpt, runPpt, supportsPpt } from '../run.ts'
import type { FormatPart, ShapeFormat } from './painterStore.ts'
import { loadFormat, loadParts, loadStyles, saveFormat } from './painterStore.ts'

/** Shapes with a fill and a text frame. */
const FILLED = new Set(['GeometricShape', 'TextBox', 'Placeholder', 'Callout'])
/** Shapes with an outline. */
const OUTLINED = new Set([...FILLED, 'Image', 'Line'])

const hex = (colour: string): string => colour.replace(/^#/, '').toUpperCase()
const plural = (n: number, word: string): string => `${n} ${word}${n === 1 ? '' : 's'}`

/** Reads the format of the one selected shape and keeps it for applying later. */
export async function pickUpFormat(): Promise<string> {
  requirePpt('1.5', 'The format painter')
  const newer = supportsPpt('1.10')
  const format = await runPpt(async (context): Promise<ShapeFormat> => {
    const shapes = context.presentation.getSelectedShapes()
    shapes.load(`items/id,items/name,items/type,items/left,items/top,items/width,items/height${newer ? ',items/rotation' : ''}`)
    await context.sync()
    const [shape] = shapes.items
    if (shapes.items.length !== 1 || !shape) throw new UserFacingError('Select one shape to pick up its format.')
    const fill = FILLED.has(shape.type) ? shape.fill.load('type,foregroundColor,transparency') : null
    const line = OUTLINED.has(shape.type) ? shape.lineFormat.load('visible,color,weight,dashStyle,transparency') : null
    const frame = FILLED.has(shape.type) ? shape.textFrame.load('leftMargin,rightMargin,topMargin,bottomMargin,autoSizeSetting,verticalAlignment,wordWrap') : null
    const range = frame ? shape.textFrame.textRange.load('text') : null
    const adjustments = newer && shape.type === 'GeometricShape' ? shape.adjustments.load('count') : null
    await context.sync()
    // Mixed text: the first character's font stands for the whole, so what you see first is what you copy.
    const font = range ? (range.text.length ? range.getSubstring(0, 1) : range).font.load('name,size,color,bold,italic') : null
    const values = adjustments ? Array.from({ length: adjustments.count }, (_, i) => adjustments.get(i)) : []
    if (font || values.length) await context.sync()
    return {
      version: 1,
      source: shape.name,
      type: shape.type,
      position: { left: shape.left, top: shape.top },
      size: { width: shape.width, height: shape.height },
      rotation: newer ? shape.rotation : null,
      fill: !fill ? null : fill.type === 'Solid' ? { kind: 'solid', colour: hex(fill.foregroundColor), transparency: fill.transparency } : fill.type === 'NoFill' ? { kind: 'none' } : null,
      line: line ? { visible: line.visible, colour: hex(line.color), weight: line.weight, dashStyle: line.dashStyle, transparency: line.transparency } : null,
      text: font ? { name: font.name, size: font.size, colour: font.color ? hex(font.color) : null, bold: font.bold, italic: font.italic } : null,
      frame: frame ? {
        margins: { left: frame.leftMargin, right: frame.rightMargin, top: frame.topMargin, bottom: frame.bottomMargin },
        autoSize: frame.autoSizeSetting, verticalAlignment: frame.verticalAlignment, wordWrap: frame.wordWrap,
      } : null,
      adjustments: adjustments ? values.map((v) => v.value) : null,
    }
  })
  saveFormat(format)
  const gaps = format.fill === null && FILLED.has(format.type) ? ' Its fill is a gradient, pattern or picture, which the painter cannot copy.' : ''
  return `Format picked up from “${format.source}”. Select shapes on any slide, then apply it.${gaps}`
}

function applyFont(font: PowerPoint.ShapeFont, f: NonNullable<ShapeFormat['text']>): void {
  if (f.name !== null) font.name = f.name
  if (f.size !== null) font.size = f.size
  if (f.colour !== null) font.color = `#${f.colour}`
  if (f.bold !== null) font.bold = f.bold
  if (f.italic !== null) font.italic = f.italic
}

/** Applies the chosen parts of a format (the one picked up last, by default) to the selected shapes. */
export async function applyFormat(parts: readonly FormatPart[] = loadParts(), format: ShapeFormat | null = loadFormat()): Promise<string> {
  if (!format) throw new UserFacingError('Pick up a format first: select one shape and choose Pick up format.')
  if (!parts.length) throw new UserFacingError('Choose what to apply first.')
  requirePpt('1.5', 'The format painter')
  const want = new Set(parts)
  const newer = supportsPpt('1.10')
  return runPpt(async (context) => {
    const shapes = context.presentation.getSelectedShapes()
    shapes.load('items/id,items/type')
    await context.sync()
    if (!shapes.items.length) throw new UserFacingError('Select the shapes to apply the format to.')
    const notes = new Set<string>()
    const wanted = format.adjustments
    const adjustable = want.has('adjustments') && newer && wanted?.length ? shapes.items.filter((s) => s.type === 'GeometricShape') : []
    const counts = new Map(adjustable.map((s) => [s, s.adjustments.load('count')]))
    if (counts.size) await context.sync()
    let changed = 0
    for (const s of shapes.items) {
      let did = false
      if (want.has('position')) {
        s.left = format.position.left
        s.top = format.position.top
        did = true
      }
      if (want.has('size')) {
        s.width = format.size.width
        s.height = format.size.height
        did = true
      }
      if (want.has('rotation')) {
        if (newer && format.rotation !== null) {
          s.rotation = format.rotation
          did = true
        } else notes.add('Rotation needs a newer PowerPoint, so it was left as it was.')
      }
      if (want.has('fillLine') && format.fill && FILLED.has(s.type)) {
        if (format.fill.kind === 'none') s.fill.clear()
        else {
          s.fill.setSolidColor(`#${format.fill.colour}`)
          s.fill.transparency = format.fill.transparency
        }
        did = true
      }
      if (want.has('fillLine') && format.line && OUTLINED.has(s.type)) {
        const l = format.line
        s.lineFormat.visible = l.visible
        if (l.visible) {
          s.lineFormat.color = `#${l.colour}`
          s.lineFormat.weight = l.weight
          s.lineFormat.dashStyle = l.dashStyle as PowerPoint.ShapeLineDashStyle
          s.lineFormat.transparency = l.transparency
        }
        did = true
      }
      if (want.has('text') && format.text && FILLED.has(s.type)) {
        applyFont(s.textFrame.textRange.font, format.text)
        did = true
      }
      if (want.has('frame') && format.frame && FILLED.has(s.type)) {
        const f = format.frame
        const tf = s.textFrame
        tf.leftMargin = f.margins.left
        tf.rightMargin = f.margins.right
        tf.topMargin = f.margins.top
        tf.bottomMargin = f.margins.bottom
        tf.autoSizeSetting = f.autoSize as PowerPoint.ShapeAutoSize
        tf.verticalAlignment = f.verticalAlignment as PowerPoint.TextVerticalAlignment
        tf.wordWrap = f.wordWrap
        did = true
      }
      const count = counts.get(s)
      if (count && wanted) {
        if (count.count === wanted.length) {
          wanted.forEach((v, i) => s.adjustments.set(i, v))
          did = true
        } else notes.add('Adjustments go only to shapes of the same kind, so some were left as they were.')
      }
      if (did) changed += 1
    }
    if (want.has('adjustments') && !newer) notes.add('Shape adjustments need a newer PowerPoint.')
    await context.sync()
    const done = changed
      ? `Format from “${format.source}” applied to ${plural(changed, 'shape')}.`
      : 'Nothing to apply: the selected shapes have none of the parts you chose.'
    return [done, ...notes].join(' ')
  })
}

/** Applies a saved style, with the parts chosen in the pane. */
export async function applyNamedStyle(id: string, parts: readonly FormatPart[] = loadParts()): Promise<string> {
  const style = loadStyles().find((s) => s.id === id)
  if (!style) throw new UserFacingError('That style is no longer saved.')
  return applyFormat(parts, style.format)
}
