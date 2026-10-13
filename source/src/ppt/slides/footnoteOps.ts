// Numbered footnotes in PowerPoint: a superscript marker at the cursor, the footnote block above the source line,
// renumbering in reading order and removing notes no marker refers to. The planning is in footnotes.ts.

import { UserFacingError } from '../../word/run.ts'
import { hasTag } from '../check/shapes.ts'
import { takeDeckSnapshot } from '../check/snapshot.ts'
import { oneLine } from '../check/text.ts'
import type { ShapeInfo, SlideInfo } from '../check/types.ts'
import { plural } from '../check/util.ts'
import { requirePpt, runPpt } from '../run.ts'
import type { Footnote, PlacedMarker } from './footnotes.ts'
import { FOOTNOTES_TAG, FOOTNOTE_STYLE, footnoteBox, formatFootnotes, nextFootnoteNumber, parseFootnotes, renumberPlan, unusedFootnotes } from './footnotes.ts'
import { shapeAt } from './office.ts'

/** Markers are superscript, which PowerPoint reads and writes from API 1.8. */
export const FOOTNOTES_NEED = '1.8'

interface SlideNotes {
  slide: SlideInfo
  block: ShapeInfo | null
  notes: Footnote[]
  markers: PlacedMarker[]
}

export function notesOf(slide: SlideInfo): SlideNotes {
  const block = slide.shapes.find((s) => hasTag(s, FOOTNOTES_TAG)) ?? null
  return {
    slide,
    block,
    notes: block ? parseFootnotes(block.text ?? '') : [],
    markers: slide.shapes.flatMap((s) => s.markers.map((m) => ({ ...m, path: s.path, top: s.top, left: s.left }))),
  }
}

async function readSlide(slideId?: string): Promise<SlideNotes> {
  const snapshot = await takeDeckSnapshot({ formatting: false, markers: true, ...(slideId ? { slideIds: [slideId] } : { current: true }) })
  const slide = snapshot.slides[0]
  if (!slide) throw new UserFacingError('Select a slide first, in the thumbnails on the left.')
  return notesOf(slide)
}

/** Writes the notes into the block: a new one above the source line, or the old one resized from its bottom edge. */
function queueBlock(target: PowerPoint.Slide, current: SlideNotes, notes: Footnote[]): void {
  const { block } = current
  if (!notes.length) {
    if (block) shapeAt(target, block.path).delete()
    return
  }
  const box = footnoteBox(notes)
  if (block) {
    const shape = shapeAt(target, block.path)
    shape.textFrame.textRange.text = formatFootnotes(notes)
    shape.top = block.top + block.height - box.height
    shape.height = box.height
    return
  }
  const shape = target.shapes.addTextBox(formatFootnotes(notes), { ...box })
  shape.name = 'BwB footnotes'
  shape.tags.add(FOOTNOTES_TAG, 'notes')
  const frame = shape.textFrame
  frame.autoSizeSetting = 'AutoSizeNone'
  frame.wordWrap = true
  frame.verticalAlignment = 'Bottom'
  frame.leftMargin = 0
  frame.rightMargin = 0
  frame.topMargin = 0
  frame.bottomMargin = 0
  const font = frame.textRange.font
  font.name = FOOTNOTE_STYLE.font
  font.size = FOOTNOTE_STYLE.size
  font.color = `#${FOOTNOTE_STYLE.colour}`
  font.bold = false
}

/** Where the text cursor is: the slide, the shape and the position after any selected text. */
async function cursor(): Promise<{ slideId: string; shapeId: string; position: number }> {
  return runPpt(async (context) => {
    const range = context.presentation.getSelectedTextRangeOrNullObject()
    range.load('start,length')
    await context.sync()
    if (range.isNullObject) throw new UserFacingError('Click in the text where the footnote marker goes, then try again.')
    const shape = range.getParentTextFrame().getParentShape()
    shape.load('id')
    const slides = context.presentation.getSelectedSlides()
    slides.load('items/id')
    await context.sync()
    const slideId = slides.items[0]?.id
    if (!slideId) throw new UserFacingError('Click in the text where the footnote marker goes, then try again.')
    return { slideId, shapeId: shape.id, position: range.start + range.length }
  })
}

/** Puts the next number at the cursor as a superscript marker, and the note in the slide's footnote block. */
export async function addFootnote(note: string): Promise<string> {
  const text = oneLine(note)
  if (!text) throw new UserFacingError('Type the footnote first.')
  requirePpt(FOOTNOTES_NEED, 'Footnote markers')
  const at = await cursor()
  const current = await readSlide(at.slideId)
  const shape = current.slide.shapes.find((s) => s.id === at.shapeId)
  if (!shape) throw new UserFacingError('Click in the text where the footnote marker goes, then try again.')
  if (hasTag(shape, FOOTNOTES_TAG)) throw new UserFacingError('Click in the slide’s text where the marker goes, not in the footnotes.')
  const number = nextFootnoteNumber(current.markers.map((m) => m.number), current.notes)
  const marker = String(number)
  await runPpt(async (context) => {
    const target = context.presentation.slides.getItem(at.slideId)
    const range = shapeAt(target, shape.path).textFrame.textRange
    range.getSubstring(at.position, 0).text = marker
    range.getSubstring(at.position, marker.length).font.superscript = true
    queueBlock(target, current, [...current.notes, { number, text }])
    await context.sync()
  })
  return `Footnote ${number} added.`
}

/** Numbers the markers 1, 2, 3 in reading order and puts the notes in the same order. */
export async function renumberFootnotes(slideId?: string): Promise<string> {
  requirePpt(FOOTNOTES_NEED, 'Renumbering footnotes')
  const current = await readSlide(slideId)
  if (!current.markers.length && !current.notes.length) return 'This slide has no footnotes.'
  const plan = renumberPlan(current.markers, current.notes)
  if (!plan.changes.length && formatFootnotes(plan.notes) === formatFootnotes(current.notes)) return 'The footnotes are already in order.'
  await runPpt(async (context) => {
    const target = context.presentation.slides.getItem(current.slide.id)
    const byShape = new Map<string, typeof plan.changes>()
    for (const c of plan.changes) byShape.set(c.path.join('/'), [...(byShape.get(c.path.join('/')) ?? []), c])
    for (const changes of byShape.values()) {
      const first = changes[0]
      if (!first) continue
      const range = shapeAt(target, first.path).textFrame.textRange
      // From the end of the text back, so a longer or shorter number does not move the ones before it.
      for (const c of [...changes].sort((a, b) => b.start - a.start)) range.getSubstring(c.start, c.length).text = c.text
    }
    queueBlock(target, current, plan.notes)
    await context.sync()
  })
  return plan.mapping.size ? `Footnotes renumbered 1 to ${plan.mapping.size}.` : 'Footnotes renumbered.'
}

/** Removes the notes no marker refers to; returns how many. */
export async function dropUnusedFootnotes(slideId?: string): Promise<number> {
  requirePpt(FOOTNOTES_NEED, 'Removing unused footnotes')
  const current = await readSlide(slideId)
  const unused = new Set(unusedFootnotes(current.markers.map((m) => m.number), current.notes))
  if (!unused.size) return 0
  await runPpt(async (context) => {
    queueBlock(context.presentation.slides.getItem(current.slide.id), current, current.notes.filter((n) => !unused.has(n)))
    await context.sync()
  })
  return unused.size
}

export async function removeUnusedFootnotes(): Promise<string> {
  const n = await dropUnusedFootnotes()
  return n ? `${plural(n, 'footnote')} with no marker removed. Renumber if the numbers now have gaps.` : 'Every footnote on this slide has a marker.'
}
