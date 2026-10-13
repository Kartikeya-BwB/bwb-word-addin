// Layout: shapes off the slide, empty placeholders and text boxes, and titles that jump from slide to slide.

import { TITLE_AREA } from '../../brand/slideGrid.ts'
import {
  FOOTER_PLACEHOLDERS, NON_TEXT_CONTENT, TEXT_PLACEHOLDERS, TITLE_PLACEHOLDERS, hasText, isBwbFurniture, isEmptySlide, isPicture, slideKind,
} from './shapes.ts'
import { points } from './text.ts'
import type { Box, DeckFinding, DeckSnapshot, ShapeInfo, SlideInfo } from './types.ts'
import { onShape, shapeExcerpt, shapeKey } from './util.ts'

/** Slack for rounding: PowerPoint reports positions to fractions of a point. */
const EDGE = 0.5
/** Titles within this many points of the usual position and size count as lined up. */
export const TITLE_TOLERANCE = 2

export const boxOf = (s: Box): Box => ({ left: s.left, top: s.top, width: s.width, height: s.height })

function offSlideFindings(snapshot: DeckSnapshot): DeckFinding[] {
  const out: DeckFinding[] = []
  const { width: W, height: H } = snapshot
  for (const slide of snapshot.slides) {
    for (const shape of slide.shapes.filter((s) => s.level === 0)) {
      const right = shape.left + shape.width
      const bottom = shape.top + shape.height
      const wholly = right <= 0 || bottom <= 0 || shape.left >= W || shape.top >= H
      const partly = !wholly && (shape.left < -EDGE || shape.top < -EDGE || right > W + EDGE || bottom > H + EDGE)
      const where = { slide: slide.index + 1, location: onShape(slide, shape), excerpt: shapeExcerpt(shape) }
      const ref = { slideId: slide.id, path: shape.path, from: boxOf(shape) }
      if (wholly) {
        out.push({
          id: `layout:off:${shapeKey(slide, shape)}`, rule: 'layout/off-slide', group: 'layout', severity: 'minor',
          title: 'Shape parked off the slide',
          detail: 'Nobody sees it in the show, but it travels with the file to the client. Delete it, or move it back if it belongs.',
          ...where, fix: { ...ref, kind: 'delete-shape' }, fixLabel: 'Delete it',
        })
      } else if (partly) {
        const fits = shape.width <= W && shape.height <= H
        const to = {
          left: Math.min(Math.max(shape.left, 0), W - shape.width),
          top: Math.min(Math.max(shape.top, 0), H - shape.height),
          width: shape.width,
          height: shape.height,
        }
        out.push({
          id: `layout:edge:${shapeKey(slide, shape)}`, rule: 'layout/off-edge', group: 'layout', severity: 'minor',
          // A photo that bleeds off the edge is often meant to.
          ...(isPicture(shape) ? { advisory: true } : {}),
          title: 'Shape runs off the edge of the slide', detail: 'The part beyond the edge is cut off in the show and in PDFs.',
          ...where, ...(fits ? { fix: { ...ref, kind: 'move-onto-slide' as const, to }, fixLabel: 'Move it onto the slide' } : {}),
        })
      }
    }
  }
  return out
}

function emptyFindings(snapshot: DeckSnapshot): DeckFinding[] {
  const out: DeckFinding[] = []
  for (const slide of snapshot.slides) {
    // An empty slide is one finding (in the slide rules), not one for each of its placeholders.
    if (isEmptySlide(slide, snapshot.can.placeholders)) continue
    for (const shape of slide.shapes) {
      const where = { slide: slide.index + 1, location: onShape(slide, shape), excerpt: shape.name }
      const ref = { slideId: slide.id, path: shape.path }
      const p = shape.placeholder
      if (shape.type === 'Placeholder' && p !== null && !FOOTER_PLACEHOLDERS.has(p) && !TITLE_PLACEHOLDERS.has(p)) {
        const text = TEXT_PLACEHOLDERS.has(p)
        const holdsContent = shape.contains !== null && NON_TEXT_CONTENT.has(shape.contains)
        // Empty text placeholders have no text; picture, chart and table placeholders hold nothing (API 1.8 says so).
        const empty = text ? shape.text !== null && !hasText(shape) && !holdsContent : snapshot.can.placeholders && p !== 'Unsupported' && shape.contains === null
        if (empty) {
          out.push({
            id: `layout:placeholder:${shapeKey(slide, shape)}`, rule: 'layout/empty-placeholder', group: 'layout', severity: 'minor',
            title: 'Empty placeholder', detail: 'PowerPoint shows its “Click to add” prompt to anyone who edits the deck. Fill it in or delete it.',
            ...where, fix: { ...ref, kind: 'delete-empty', textual: text }, fixLabel: 'Delete it',
          })
        }
      }
      const filled = shape.fill !== null && shape.fill.type !== 'NoFill'
      const outlined = shape.line !== null && shape.line.visible
      if (shape.type === 'TextBox' && shape.text !== null && !hasText(shape) && shape.fill !== null && !filled && !outlined && !isBwbFurniture(shape)) {
        out.push({
          id: `layout:textbox:${shapeKey(slide, shape)}`, rule: 'layout/empty-text-box', group: 'layout', severity: 'copyedit',
          title: 'Empty text box', detail: 'It holds nothing but gets in the way when you select or move shapes.',
          ...where, fix: { ...ref, kind: 'delete-empty', textual: true }, fixLabel: 'Delete it',
        })
      }
    }
  }
  return out
}

const near = (a: Box, b: Box): boolean =>
  Math.abs(a.left - b.left) <= TITLE_TOLERANCE && Math.abs(a.top - b.top) <= TITLE_TOLERANCE &&
  Math.abs(a.width - b.width) <= TITLE_TOLERANCE && Math.abs(a.height - b.height) <= TITLE_TOLERANCE

const gap = (a: Box, b: Box): number => Math.abs(a.left - b.left) + Math.abs(a.top - b.top) + Math.abs(a.width - b.width) + Math.abs(a.height - b.height)

/** The usual title geometry: the one most titles sit within 2 pt of; on a tie, the one closer to the BwB grid. */
export function usualTitleBox(boxes: Box[]): { box: Box; count: number } | null {
  let best: { box: Box; count: number; gap: number } | null = null
  for (const b of boxes) {
    const count = boxes.filter((o) => near(b, o)).length
    const g = gap(b, TITLE_AREA)
    if (!best || count > best.count || (count === best.count && g < best.gap)) best = { box: b, count, gap: g }
  }
  return best && best.count >= 2 ? { box: best.box, count: best.count } : null
}

function describeJump(from: Box, to: Box): string {
  const parts: string[] = []
  const dy = from.top - to.top
  const dx = from.left - to.left
  if (Math.abs(dy) > TITLE_TOLERANCE) parts.push(`${points(Math.round(Math.abs(dy) * 10) / 10)} ${dy > 0 ? 'lower' : 'higher'}`)
  if (Math.abs(dx) > TITLE_TOLERANCE) parts.push(`${points(Math.round(Math.abs(dx) * 10) / 10)} to the ${dx > 0 ? 'right' : 'left'}`)
  if (Math.abs(from.width - to.width) > TITLE_TOLERANCE || Math.abs(from.height - to.height) > TITLE_TOLERANCE) parts.push('a different size')
  return parts.join(' and ')
}

function jumpingTitleFindings(snapshot: DeckSnapshot): DeckFinding[] {
  // Title, divider and closing slides place their titles differently by design.
  const titles: Array<{ slide: SlideInfo; shape: ShapeInfo }> = snapshot.slides
    .filter((slide) => !['title', 'divider', 'closing'].includes(slideKind(slide)))
    .flatMap((slide) => slide.shapes.filter((s) => s.placeholder === 'Title' && s.level === 0).map((shape) => ({ slide, shape })))
  if (titles.length < 3) return []
  const usual = usualTitleBox(titles.map((t) => boxOf(t.shape)))
  if (!usual) return []
  return titles
    .filter((t) => !near(boxOf(t.shape), usual.box))
    .map(({ slide, shape }) => ({
      id: `layout:title:${shapeKey(slide, shape)}`, rule: 'layout/jumping-title', group: 'layout', severity: 'minor',
      title: 'Title out of line with the other slides',
      detail: `Compared with the title on ${usual.count} other slides, it is ${describeJump(boxOf(shape), usual.box)}. Titles that jump as you click through look careless.`,
      excerpt: shapeExcerpt(shape), slide: slide.index + 1, location: onShape(slide, shape),
      fix: { kind: 'align-title', slideId: slide.id, path: shape.path, from: boxOf(shape), to: usual.box },
      fixLabel: 'Line it up with the others',
    }))
}

export function layoutFindings(snapshot: DeckSnapshot): DeckFinding[] {
  return [...offSlideFindings(snapshot), ...emptyFindings(snapshot), ...jumpingTitleFindings(snapshot)]
}

