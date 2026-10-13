// What the Slides tab shows about the deck, from one read: the storyline, the sections and the state of the trackers.

import { takeDeckSnapshot } from '../check/snapshot.ts'
import type { DeckSnapshot } from '../check/types.ts'
import type { Section, TrackerIssue } from './sections.ts'
import { findSections, trackerIssues, trackersOn } from './sections.ts'
import type { StoryEntry } from './storyline.ts'
import { storyline } from './storyline.ts'

export interface DeckOutline {
  entries: StoryEntry[]
  sections: Section[]
  /** How many slides carry a tracker, and which slides' trackers no longer match the sections. */
  trackers: { slides: number; issues: TrackerIssue[] }
  /** The current slide's id, when PowerPoint can say. */
  current: string | null
}

export function outlineOf(snapshot: DeckSnapshot): DeckOutline {
  const sections = findSections(snapshot.slides)
  return {
    entries: storyline(snapshot.slides),
    sections,
    trackers: {
      slides: snapshot.slides.filter((s) => trackersOn(s).length).length,
      issues: trackerIssues(snapshot.slides, sections),
    },
    current: snapshot.selected[0] ?? null,
  }
}

export async function readOutline(): Promise<DeckOutline> {
  return outlineOf(await takeDeckSnapshot({ formatting: false }))
}
