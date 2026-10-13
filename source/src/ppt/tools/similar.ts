// Select similar: from one selected shape, selects every shape on the slide like it, by fill colour, size (to half a
// point), type or the font of its text. Needs PowerPointApi 1.5 to read and set the selection.

import { UserFacingError } from '../../word/run.ts'
import { currentSlide } from '../render.ts'
import { requirePpt, runPpt } from '../run.ts'

export type Likeness = 'fill' | 'size' | 'type' | 'font'

export const LIKENESSES: Array<{ id: Likeness; label: string; done: string }> = [
  { id: 'fill', label: 'Same fill', done: 'the same fill' },
  { id: 'size', label: 'Same size', done: 'the same size' },
  { id: 'type', label: 'Same type', done: 'the same type' },
  { id: 'font', label: 'Same font', done: 'the same font' },
]

/** Shapes with a fill and text. */
const FILLED = new Set(['GeometricShape', 'TextBox', 'Placeholder', 'Callout'])
/** Sizes within this many points count as the same. */
export const SIZE_TOLERANCE = 0.5

/** What two fills share when they look the same: none, a solid colour, or the same kind of other fill. */
export function fillKey(type: string, colour: string): string {
  if (type === 'NoFill') return 'none'
  return type === 'Solid' ? `solid:${colour.replace(/^#/, '').toUpperCase()}` : `other:${type}`
}

export function sameSize(a: { width: number; height: number }, b: { width: number; height: number }): boolean {
  return Math.abs(a.width - b.width) <= SIZE_TOLERANCE && Math.abs(a.height - b.height) <= SIZE_TOLERANCE
}

interface Probe {
  shape: PowerPoint.Shape
  fill: PowerPoint.ShapeFill | null
  text: PowerPoint.TextRange | null
}

/** Loads what the comparison needs for each shape: fills and text only where the shape has them. */
function probe(shape: PowerPoint.Shape, by: Likeness): Probe {
  const filled = FILLED.has(shape.type)
  return {
    shape,
    fill: by === 'fill' && filled ? shape.fill.load('type,foregroundColor') : null,
    text: by === 'font' && filled ? shape.textFrame.textRange.load('text') : null,
  }
}

/** The first character's font name: when the text is mixed, what comes first stands for the rest. */
function firstFont(text: PowerPoint.TextRange): PowerPoint.ShapeFont {
  return (text.text.length ? text.getSubstring(0, 1) : text).font.load('name')
}

/** Selects every shape on the slide like the selected one. */
export async function selectSimilar(by: Likeness): Promise<string> {
  requirePpt('1.5', 'Select similar')
  return runPpt(async (context) => {
    const selected = context.presentation.getSelectedShapes()
    selected.load('items/id,items/type,items/width,items/height')
    const slide = await currentSlide(context)
    const all = slide.shapes
    all.load('items/id,items/type,items/width,items/height')
    await context.sync()
    const [ref] = selected.items
    if (selected.items.length !== 1 || !ref) throw new UserFacingError('Select one shape: the others on the slide are matched to it.')
    if ((by === 'fill' || by === 'font') && !FILLED.has(ref.type)) {
      throw new UserFacingError(by === 'fill' ? 'Pictures, lines, tables and groups have no fill to match.' : 'This shape holds no text, so there is no font to match.')
    }
    const mine = probe(ref, by)
    const others = all.items.map((s) => probe(s, by))
    await context.sync()
    const fonts = by === 'font' ? new Map([mine, ...others].filter((p) => p.text).map((p) => [p, firstFont(p.text as PowerPoint.TextRange)])) : null
    if (fonts) await context.sync()
    const key = (p: Probe): string | null => {
      switch (by) {
        case 'fill': return p.fill ? fillKey(p.fill.type, p.fill.foregroundColor) : null
        case 'type': return p.shape.type
        case 'font': return fonts?.get(p)?.name ?? null
        case 'size': return null
      }
    }
    const target = key(mine)
    const matches = others.filter((p) => (by === 'size' ? sameSize(p.shape, ref) : target !== null && key(p) === target)).map((p) => p.shape.id)
    const ids = matches.includes(ref.id) ? matches : [ref.id, ...matches]
    slide.setSelectedShapes(ids)
    await context.sync()
    const done = LIKENESSES.find((l) => l.id === by)?.done ?? 'the same look'
    return ids.length > 1 ? `Selected ${ids.length} shapes with ${done}.` : `No other shape on this slide has ${done}.`
  })
}
