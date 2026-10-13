// Sections, the agenda and section trackers, planned from a snapshot of the deck. A section starts at a slide tagged
// BWB_SECTION=<name> (Start a section here) or at a BwB divider slide, which is named after its title. Trackers sit
// in a row above the title of every slide from the first section on, except title, agenda, divider, summary and
// closing slides, and list the sections with the current one in navy bold.

import { TAGS, TITLE_AREA } from '../../brand/slideGrid.ts'
import { COLOURS } from '../../brand/tokens.ts'
import type { SlideKind } from '../check/shapes.ts'
import { hasTag, slideKind, slideTitle, tagValue } from '../check/shapes.ts'
import { oneLine, shorten } from '../check/text.ts'
import type { Box, ShapeInfo, SlideInfo, Span } from '../check/types.ts'

export interface Section {
  name: string
  slideId: string
  /** 0-based position of the slide that starts it. */
  index: number
  /** A slide marked with Start a section here, or a divider slide. */
  from: 'tag' | 'divider'
}

export const TRACKER_SEPARATOR = ' | '
/** A row above the title, clear of stickers at the top right. */
export const TRACKER_BOX: Box = { left: TITLE_AREA.left, top: 8, width: 640, height: 16 }
export const TRACKER_STYLE = { font: 'Aptos', size: 9, colour: COLOURS.charcoal, current: COLOURS.navy } as const
export const TRACKER_NAME_MAX = 32

/** Slides that never carry a tracker. */
const NO_TRACKER = new Set<SlideKind>(['title', 'agenda', 'divider', 'summary', 'closing'])

export function findSections(slides: SlideInfo[]): Section[] {
  const out: Section[] = []
  for (const slide of slides) {
    const tagged = oneLine(tagValue(slide, TAGS.section) ?? '')
    if (tagged) {
      out.push({ name: tagged, slideId: slide.id, index: slide.index, from: 'tag' })
    } else if (slideKind(slide) === 'divider') {
      out.push({ name: slideTitle(slide)?.text || `Section ${out.length + 1}`, slideId: slide.id, index: slide.index, from: 'divider' })
    }
  }
  return out
}

/** The agenda's lines: the sections, numbered. */
export function agendaLines(sections: Section[]): string[] {
  return sections.map((s, i) => `${i + 1}. ${s.name}`)
}

/** The section a slide belongs to: the last one that starts at or before it, or -1 before the first. */
export function sectionOf(sections: Section[], index: number): number {
  let found = -1
  sections.forEach((s, i) => {
    if (s.index <= index) found = i
  })
  return found
}

export interface TrackerText {
  text: string
  /** Where the current section's name is, to set in navy bold. */
  current: Span
}

export function trackerText(sections: Section[], current: number): TrackerText {
  let start = 0
  let span: Span = { start: 0, length: 0 }
  const names = sections.map((s, i) => {
    const name = shorten(s.name, TRACKER_NAME_MAX)
    if (i === current) span = { start, length: name.length }
    start += name.length + TRACKER_SEPARATOR.length
    return name
  })
  return { text: names.join(TRACKER_SEPARATOR), current: span }
}

export interface TrackerPlan extends TrackerText {
  slideId: string
  index: number
  section: number
  /** Written to the tracker's BWB_TRACKER tag: the section it highlights, so a moved slide shows as out of date. */
  value: string
}

export function trackerValue(sections: Section[], section: number): string {
  return `${section + 1}:${sections[section]?.name ?? ''}`
}

/** The tracker every slide should carry, given the sections. */
export function planTrackers(slides: SlideInfo[], sections: Section[] = findSections(slides)): TrackerPlan[] {
  const first = sections[0]
  if (!first) return []
  return slides.flatMap((slide) => {
    if (slide.index < first.index || NO_TRACKER.has(slideKind(slide))) return []
    const section = sectionOf(sections, slide.index)
    return [{ slideId: slide.id, index: slide.index, section, value: trackerValue(sections, section), ...trackerText(sections, section) }]
  })
}

export function trackersOn(slide: SlideInfo): ShapeInfo[] {
  return slide.shapes.filter((s) => hasTag(s, TAGS.tracker))
}

export interface TrackerIssue {
  slideId: string
  index: number
  problem: 'missing' | 'outdated' | 'extra'
}

/** Slides whose trackers no longer match the sections. A deck without trackers has nothing to keep up to date. */
export function trackerIssues(slides: SlideInfo[], sections: Section[] = findSections(slides)): TrackerIssue[] {
  if (!slides.some((s) => trackersOn(s).length)) return []
  const plans = new Map(planTrackers(slides, sections).map((p) => [p.slideId, p]))
  const out: TrackerIssue[] = []
  for (const slide of slides) {
    const trackers = trackersOn(slide)
    const plan = plans.get(slide.id)
    const at = { slideId: slide.id, index: slide.index }
    if (!plan) {
      if (trackers.length) out.push({ ...at, problem: 'extra' })
    } else if (!trackers.length) {
      out.push({ ...at, problem: 'missing' })
    } else {
      const fresh = trackers.length === 1 && trackers.every((t) => oneLine(t.text ?? '') === plan.text && tagValue(t, TAGS.tracker) === plan.value)
      if (!fresh) out.push({ ...at, problem: 'outdated' })
    }
  }
  return out
}
