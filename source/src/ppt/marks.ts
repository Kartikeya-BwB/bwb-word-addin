// Marks on slides: status stickers, review notes, Harvey balls, status marks and RAG dots.

import type { IconIndex } from '../brand/icons.ts'
import { ICON_SIZES_CM, loadIconIndex } from '../brand/icons.ts'
import { SLIDE, TAGS } from '../brand/slideGrid.ts'
import type { Rect } from '../geometry/index.ts'
import { cmToPt } from '../geometry/index.ts'
import { insertSvgHere } from '../media/insert.ts'
import type { Sticker } from '../word/stickers.ts'
import { UserFacingError } from '../word/run.ts'
import { addShape, currentSlide, insertElement, taggedShapes } from './render.ts'
import { requirePpt, runPpt, supportsPpt } from './run.ts'
import type { RagColour } from './specs.ts'
import { markPosition, noteSpec, ragSpec, stickerSpec } from './specs.ts'
import type { Quarters, StatusKind } from './svg.ts'
import { HARVEY_LABELS, STATUS_MARKS, harveySvg, statusSvg } from './svg.ts'

const plural = (n: number, word: string): string => `${n} ${word}${n === 1 ? '' : 's'}`

const MONTHS = ['January', 'February', 'March', 'April', 'May', 'June', 'July', 'August', 'September', 'October', 'November', 'December']

/** A date as BwB writes it: 13 October 2026. */
export function formatDate(d: Date): string {
  return `${d.getDate()} ${MONTHS[d.getMonth()] ?? ''} ${d.getFullYear()}`
}

export function isoDate(d: Date): string {
  return `${d.getFullYear()}-${String(d.getMonth() + 1).padStart(2, '0')}-${String(d.getDate()).padStart(2, '0')}`
}

/** Puts a sticker on each selected slide, replacing the one there. */
export async function setSlideStickers(label: Sticker): Promise<string> {
  requirePpt('1.5', 'Adding stickers')
  const wide = supportsPpt('1.10')
  return runPpt(async (context) => {
    const slides = context.presentation.getSelectedSlides()
    slides.load('items/id')
    const page = wide ? context.presentation.pageSetup : null
    page?.load('slideWidth')
    await context.sync()
    if (!slides.items.length) throw new UserFacingError('Select the slides to mark first.')
    for (const old of await taggedShapes(context, slides.items, TAGS.sticker)) old.shape.delete()
    const spec = stickerSpec(label, page?.slideWidth ?? SLIDE.width)
    for (const slide of slides.items) {
      const shape = addShape(slide.shapes, spec, false)
      shape.tags.add(TAGS.sticker, label)
      shape.tags.add(TAGS.element, 'sticker')
    }
    await context.sync()
    return `“${label}” sticker on ${plural(slides.items.length, 'slide')}.`
  })
}

async function removeEverywhere(key: string): Promise<number> {
  requirePpt('1.5', 'Removing marks')
  return runPpt(async (context) => {
    const slides = context.presentation.slides
    slides.load('items/id')
    await context.sync()
    const found = await taggedShapes(context, slides.items, key)
    found.forEach((t) => t.shape.delete())
    await context.sync()
    return found.length
  })
}

export async function removeAllStickers(): Promise<string> {
  const n = await removeEverywhere(TAGS.sticker)
  return n ? `Removed ${plural(n, 'sticker')}.` : 'There are no stickers in this presentation.'
}

/** A navy review note at the top of the slide with today's date, tagged so it can be removed before sending. */
export async function addReviewNote(today: Date = new Date()): Promise<string> {
  requirePpt('1.5', 'Adding review notes')
  const existing = await runPpt(async (context) => (await taggedShapes(context, [await currentSlide(context)], TAGS.note)).length)
  await insertElement(noteSpec(formatDate(today), existing), { [TAGS.note]: isoDate(today) })
  return 'Review note added. Type after “Note:”. Remove all notes before you send the deck.'
}

export async function removeAllNotes(): Promise<string> {
  const n = await removeEverywhere(TAGS.note)
  return n ? `Removed ${plural(n, 'review note')}.` : 'There are no review notes in this presentation.'
}

const MARK_KINDS = new Set(['harvey', 'status', 'rag'])

/** Where the next mark goes: beside the mark selected now, or in the middle of the content area. */
async function nextMarkPosition(size: number): Promise<{ left: number; top: number }> {
  requirePpt('1.5', 'Inserting marks')
  return runPpt(async (context) => {
    const shapes = context.presentation.getSelectedShapes()
    shapes.load('items/left,items/top,items/width,items/height')
    await context.sync()
    const only = shapes.items.length === 1 ? shapes.items[0] : undefined
    let previous: Rect | null = null
    if (only) {
      const tag = only.tags.getItemOrNullObject(TAGS.element)
      tag.load('value')
      await context.sync()
      if (!tag.isNullObject && MARK_KINDS.has(tag.value)) previous = { left: only.left, top: only.top, width: only.width, height: only.height }
    }
    return markPosition(previous, size)
  })
}

/** Names and tags the graphic PowerPoint just inserted and selected, if the selection is indeed that graphic. */
async function tagInserted(name: string, kind: string, at: { left: number; top: number }, alt: string): Promise<void> {
  const described = supportsPpt('1.10')
  await runPpt(async (context) => {
    const shapes = context.presentation.getSelectedShapes()
    shapes.load('items/left,items/top')
    await context.sync()
    const [shape] = shapes.items
    if (shapes.items.length !== 1 || !shape || Math.abs(shape.left - at.left) > 1 || Math.abs(shape.top - at.top) > 1) return
    shape.name = name
    shape.tags.add(TAGS.element, kind)
    if (described) shape.altTextDescription = alt
    await context.sync()
  })
}

const MARK_SIZE = cmToPt(ICON_SIZES_CM.marker)

export async function insertHarveyBall(quarters: Quarters): Promise<string> {
  const at = await nextMarkPosition(MARK_SIZE)
  await insertSvgHere({ svg: harveySvg(quarters), widthPt: MARK_SIZE, heightPt: MARK_SIZE, altText: HARVEY_LABELS[quarters], ...at })
  await tagInserted(`BwB Harvey ball ${quarters}`, 'harvey', at, HARVEY_LABELS[quarters])
  return 'Harvey ball added. Insert another to put it alongside.'
}

export async function insertStatusMark(kind: StatusKind, loadIndex: () => Promise<IconIndex> = loadIconIndex): Promise<string> {
  const svg = statusSvg(await loadIndex(), kind)
  const at = await nextMarkPosition(MARK_SIZE)
  const label = STATUS_MARKS[kind].label
  await insertSvgHere({ svg, widthPt: MARK_SIZE, heightPt: MARK_SIZE, altText: label, ...at })
  await tagInserted(`BwB status ${label.toLowerCase()}`, 'status', at, label)
  return `“${label}” mark added.`
}

export async function insertRagDot(colour: RagColour): Promise<string> {
  const size = cmToPt(ICON_SIZES_CM.inline)
  const at = await nextMarkPosition(size)
  await insertElement({ kind: 'rag', name: `BwB RAG ${colour}`, group: false, shapes: [ragSpec(colour, at)] }, { [TAGS.rag]: colour })
  return colour === 'green' ? 'Green dot added.' : `${colour === 'amber' ? 'Amber' : 'Red'} dot added. It is outside the BwB palette: use it in trackers only.`
}
