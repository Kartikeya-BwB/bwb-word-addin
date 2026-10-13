// The storyline: every slide's title in order, as consultants read a deck before the detail. Titles come from the
// title placeholder, or else the top-most text on the slide. Slides without a title, and titles over two lines, are
// flagged. From the storyline the pane goes to a slide, copies the titles and builds an executive summary slide.

import type { SlideKind } from '../check/shapes.ts'
import { LONG_TITLE, slideKind, slideTitle } from '../check/shapes.ts'
import { takeDeckSnapshot } from '../check/snapshot.ts'
import type { SlideInfo } from '../check/types.ts'
import { UserFacingError } from '../../word/run.ts'
import { insertListSlide } from './listSlide.ts'
import { NEEDS } from './office.ts'
import { requirePpt, runPpt } from '../run.ts'
import { findSections, sectionOf } from './sections.ts'

export { LONG_TITLE }

export interface StoryEntry {
  slideId: string
  /** 1-based slide number. */
  number: number
  title: string | null
  kind: SlideKind
  /** True when the title is in a title placeholder rather than the top-most text. */
  fromPlaceholder: boolean
  long: boolean
  /** The section the slide is in, if the deck has sections. */
  section: string | null
}

export function storyline(slides: SlideInfo[]): StoryEntry[] {
  const sections = findSections(slides)
  return slides.map((slide) => {
    const title = slideTitle(slide)
    const section = sections[sectionOf(sections, slide.index)]
    return {
      slideId: slide.id,
      number: slide.index + 1,
      title: title?.text ?? null,
      kind: slideKind(slide),
      fromPlaceholder: title?.fromPlaceholder ?? false,
      long: (title?.text.length ?? 0) > LONG_TITLE,
      section: section?.name ?? null,
    }
  })
}

/** The storyline as plain text, one numbered title a line, with the section names as headings when there are any. */
export function storylineText(entries: StoryEntry[]): string {
  const lines: string[] = []
  let section: string | null = null
  for (const e of entries) {
    if (e.section !== null && e.section !== section) {
      if (lines.length) lines.push('')
      lines.push(e.section.toUpperCase())
    }
    section = e.section
    lines.push(`${e.number}. ${e.title ?? '(no title)'}`)
  }
  return lines.join('\n')
}

/** The executive summary's bullets: the titles of the content slides, which carry the deck's messages. */
export function executiveSummaryLines(entries: StoryEntry[]): string[] {
  return entries.flatMap((e) => (e.kind === 'content' && e.title ? [e.title] : []))
}

export interface StorylineIssues {
  untitled: number
  long: number
}

export function storylineIssues(entries: StoryEntry[]): StorylineIssues {
  return { untitled: entries.filter((e) => e.title === null).length, long: entries.filter((e) => e.long).length }
}

export async function readStoryline(): Promise<StoryEntry[]> {
  const snapshot = await takeDeckSnapshot({ formatting: false })
  return storyline(snapshot.slides)
}

export async function goToSlide(slideId: string): Promise<void> {
  requirePpt(NEEDS.current, 'Going to a slide')
  await runPpt(async (context) => {
    context.presentation.setSelectedSlides([slideId])
    await context.sync()
  })
}

export async function insertExecutiveSummary(): Promise<string> {
  requirePpt(NEEDS.read, 'Inserting an executive summary')
  const lines = executiveSummaryLines(await readStoryline())
  if (!lines.length) throw new UserFacingError('No content slide has a title yet, so there is nothing to summarise.')
  const result = await insertListSlide({ library: ['content', 'bullets'], title: 'Executive summary', items: lines, numbered: false, layout: 'summary' })
  return `Executive summary added ${result.afterCurrent ? 'after this slide' : 'at the end of the deck'}, with ${lines.length} ${lines.length === 1 ? 'title' : 'titles'} as bullets.`
}
