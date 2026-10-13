// Before you send: review notes and stickers left on slides, and section trackers that no longer match the sections.
// Placeholder text left in ([Insert chart], XX) comes from the wording rules and lands in this group too.

import { TAGS } from '../../brand/slideGrid.ts'
import { trackerIssues } from '../slides/sections.ts'
import { hasTag, tagLabel } from './shapes.ts'
import { oneLine } from './text.ts'
import type { DeckFinding, DeckSnapshot, ShapeInfo, SlideInfo } from './types.ts'
import { excerpt, onShape, onSlide, plural } from './util.ts'

interface Placed {
  slide: SlideInfo
  shape: ShapeInfo
}

function tagged(snapshot: DeckSnapshot, tag: string): Placed[] {
  return snapshot.slides.flatMap((slide) => slide.shapes.filter((s) => hasTag(s, tag)).map((shape) => ({ slide, shape })))
}

const slidesIn = (found: Placed[]): number => new Set(found.map((f) => f.slide.id)).size

function noteFindings(snapshot: DeckSnapshot): DeckFinding[] {
  const notes = tagged(snapshot, TAGS.note)
  const first = notes[0]
  if (!first) return []
  return [{
    id: 'send:notes', rule: 'send/review-notes', group: 'send', severity: 'major',
    title: `${plural(notes.length, 'review note')} left on ${plural(slidesIn(notes), 'slide')}`,
    detail: 'Review notes are for reviewers. Remove them before the deck goes to QA or the client.',
    excerpt: excerpt(oneLine(first.shape.text ?? ''), 0, 60, 0), slide: first.slide.index + 1, location: onShape(first.slide, first.shape),
    fix: { kind: 'delete-tagged', tag: TAGS.note }, fixLabel: 'Remove all review notes',
  }]
}

function stickerFindings(snapshot: DeckSnapshot): DeckFinding[] {
  const byLabel = new Map<string, Placed[]>()
  for (const found of tagged(snapshot, TAGS.sticker)) {
    const label = tagLabel(found.shape, TAGS.sticker)
    byLabel.set(label, [...(byLabel.get(label) ?? []), found])
  }
  const out: DeckFinding[] = []
  for (const [label, found] of byLabel) {
    const first = found[0]
    if (!first) continue
    out.push({
      id: `send:sticker:${label}`, rule: 'send/sticker', group: 'send', severity: 'minor', advisory: true,
      title: label ? `“${label}” sticker on ${plural(slidesIn(found), 'slide')}` : `Sticker on ${plural(slidesIn(found), 'slide')}`,
      detail: 'Remove stickers before sending a final version, unless the label still applies.',
      slide: first.slide.index + 1, location: onShape(first.slide, first.shape),
      fix: { kind: 'delete-tagged', tag: TAGS.sticker, ...(label ? { text: label } : {}) }, fixLabel: 'Remove these stickers',
    })
  }
  return out
}

function trackerFindings(snapshot: DeckSnapshot): DeckFinding[] {
  const issues = trackerIssues(snapshot.slides)
  const first = issues[0]
  const slide = first ? snapshot.slides.find((s) => s.id === first.slideId) : undefined
  if (!first || !slide) return []
  return [{
    id: 'send:trackers', rule: 'send/trackers', group: 'send', severity: 'minor',
    title: `Section trackers out of date on ${plural(issues.length, 'slide')}`,
    detail: 'The sections or slides changed after the trackers were added. Update them so each slide shows its own section.',
    slide: slide.index + 1, location: onSlide(slide),
    fix: { kind: 'update-trackers' }, fixLabel: 'Update the trackers',
  }]
}

export function sendFindings(snapshot: DeckSnapshot): DeckFinding[] {
  return [...noteFindings(snapshot), ...stickerFindings(snapshot), ...trackerFindings(snapshot)]
}
