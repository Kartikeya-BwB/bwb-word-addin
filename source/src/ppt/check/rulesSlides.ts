// Slides: a title on every slide, in the title placeholder, on two lines at most and not the same as another slide's;
// no empty slides; not too many words on one slide; and alt text on pictures and graphics.

import type { SlideKind } from './shapes.ts'
import { LONG_TITLE, authoredText, isEmptySlide, isPicture, isTitleShape, needsAltText, slideKind, slideTitle } from './shapes.ts'
import { oneLine, wordCount } from './text.ts'
import type { DeckFinding, DeckSnapshot, SlideInfo } from './types.ts'
import { onShape, onSlide, shapeExcerpt, shapeKey } from './util.ts'

/** Slides read best with fewer words than this. */
export const MAX_WORDS = 120

/** Slides whose titles carry the argument, so two of them should never say the same thing. */
const ARGUMENT: ReadonlySet<SlideKind> = new Set<SlideKind>(['content', 'summary'])

const titleKey = (text: string): string => oneLine(text).toLowerCase().replace(/[.:;,]+$/, '')

/** The same action title on two slides: the later one is flagged, pointing at the first. */
function duplicateTitleFindings(slides: SlideInfo[]): DeckFinding[] {
  const first = new Map<string, SlideInfo>()
  const out: DeckFinding[] = []
  for (const slide of slides) {
    const title = slideTitle(slide)
    if (!title || !ARGUMENT.has(slideKind(slide))) continue
    const key = titleKey(title.text)
    const earlier = first.get(key)
    if (!earlier) {
      first.set(key, slide)
      continue
    }
    out.push({
      id: `slides:same-title:${slide.id}`, rule: 'slides/duplicate-title', group: 'slides', severity: 'minor',
      title: `Same title as slide ${earlier.index + 1}`,
      detail: 'Each slide should make its own point. If this one carries on from the other, give it a title that says what it adds.',
      excerpt: title.text, slide: slide.index + 1, location: onShape(slide, title.shape),
    })
  }
  return out
}

export function slideFindings(snapshot: DeckSnapshot): DeckFinding[] {
  const out: DeckFinding[] = duplicateTitleFindings(snapshot.slides)
  for (const slide of snapshot.slides) {
    const n = slide.index + 1
    if (isEmptySlide(slide, snapshot.can.placeholders)) {
      out.push({
        id: `slides:empty:${slide.id}`, rule: 'slides/empty-slide', group: 'slides', severity: 'major',
        title: 'Empty slide', detail: 'It has nothing on it but empty placeholders. Delete it, or fill it in.',
        slide: n, location: onSlide(slide), fix: { kind: 'delete-slide', slideId: slide.id }, fixLabel: 'Delete the slide',
      })
      continue
    }
    const title = slideTitle(slide)
    if (!title) {
      out.push({
        id: `slides:no-title:${slide.id}`, rule: 'slides/no-title', group: 'slides', severity: 'major',
        title: 'Slide without a title', detail: 'Give every slide an action title: its message in one sentence, on two lines at most.',
        slide: n, location: onSlide(slide),
      })
    } else {
      const placeholder = slide.shapes.find(isTitleShape)
      if (!title.fromPlaceholder && placeholder) {
        out.push({
          id: `slides:empty-title:${slide.id}`, rule: 'slides/empty-title', group: 'slides', severity: 'minor',
          title: 'Title placeholder left empty',
          detail: 'The title is typed in another text box. Put it in the title placeholder, so the storyline, the outline and screen readers find it.',
          excerpt: title.text, slide: n, location: onShape(slide, placeholder),
        })
      } else if (!title.fromPlaceholder) {
        out.push({
          id: `slides:title-box:${slide.id}`, rule: 'slides/title-not-placeholder', group: 'slides', severity: 'minor', advisory: true,
          title: 'Title not in a title placeholder',
          detail: 'Screen readers, the outline and PDF bookmarks look for the title placeholder. Use a layout with a title.',
          excerpt: title.text, slide: n, location: onShape(slide, title.shape),
        })
      }
      if (title.text.length > LONG_TITLE) {
        out.push({
          id: `slides:long-title:${slide.id}`, rule: 'slides/long-title', group: 'slides', severity: 'minor',
          title: `Title over two lines (${title.text.length} characters)`,
          detail: `Action titles fit on two lines at most: about ${LONG_TITLE} characters at 24 pt. Tighten the message.`,
          excerpt: title.text, slide: n, location: onShape(slide, title.shape),
        })
      }
    }

    const words = authoredText(slide).reduce((sum, s) => sum + wordCount(s.text ?? ''), 0)
    if (words > MAX_WORDS) {
      out.push({
        id: `slides:words:${slide.id}`, rule: 'slides/too-much-text', group: 'slides', severity: 'minor', advisory: true,
        title: `${words} words on the slide`,
        detail: `Slides read best with ${MAX_WORDS} words or fewer. Move detail to the speaker notes, a backup slide or the report.`,
        slide: n, location: onSlide(slide),
      })
    }

    for (const shape of slide.shapes) {
      const alt = shape.alt
      if (!needsAltText(shape) || !alt || alt.decorative || alt.title.trim() || alt.description.trim()) continue
      const what = isPicture(shape) ? 'picture' : 'graphic'
      out.push({
        id: `slides:alt:${shapeKey(slide, shape)}`, rule: 'slides/alt-text', group: 'slides', severity: 'minor',
        title: `${what === 'picture' ? 'Picture' : 'Graphic'} without alt text`,
        detail: `Describe what the ${what === 'picture' ? 'picture' : 'icon or graphic'} shows (right-click, View Alt Text), or mark it as decorative.`,
        excerpt: shapeExcerpt(shape), slide: n, location: onShape(slide, shape),
      })
    }
  }
  return out
}
