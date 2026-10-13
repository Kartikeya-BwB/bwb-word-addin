// Slides: a title on every slide, in the title placeholder and on two lines at most; not too many words on one
// slide; and alt text on pictures.

import { LONG_TITLE, authoredText, isPicture, isTitleShape, needsAltText, slideTitle } from './shapes.ts'
import { wordCount } from './text.ts'
import type { DeckFinding, DeckSnapshot } from './types.ts'
import { onShape, onSlide, shapeExcerpt, shapeKey } from './util.ts'

/** Slides read best with fewer words than this. */
export const MAX_WORDS = 120

export function slideFindings(snapshot: DeckSnapshot): DeckFinding[] {
  const out: DeckFinding[] = []
  for (const slide of snapshot.slides) {
    const n = slide.index + 1
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
