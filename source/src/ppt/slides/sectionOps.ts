// Marking sections, the agenda slide and the section trackers, in PowerPoint. The planning is in sections.ts.

import { TAGS } from '../../brand/slideGrid.ts'
import { UserFacingError } from '../../word/run.ts'
import { slideTitle } from '../check/shapes.ts'
import { takeDeckSnapshot } from '../check/snapshot.ts'
import { oneLine } from '../check/text.ts'
import { plural } from '../check/util.ts'
import { insertListSlide } from './listSlide.ts'
import { NEEDS, shapeAt } from './office.ts'
import { requirePpt, runPpt } from '../run.ts'
import type { TrackerPlan } from './sections.ts'
import { TRACKER_BOX, TRACKER_STYLE, findSections, planTrackers, trackersOn } from './sections.ts'

export class NoSectionsError extends UserFacingError {
  constructor() {
    super('The deck has no sections yet. Go to the first slide of each part and choose Start a section here, or insert BwB divider slides.')
  }
}

async function setSlideTag(slideId: string, key: string, value: string | null): Promise<void> {
  await runPpt(async (context) => {
    const tags = context.presentation.slides.getItem(slideId).tags
    if (value === null) tags.delete(key)
    else tags.add(key, value)
    await context.sync()
  })
}

/** Starts a section at the current slide, named `name` or else the slide's title. */
export async function markSectionHere(name?: string): Promise<string> {
  requirePpt(NEEDS.current, 'Starting a section at the current slide')
  const snapshot = await takeDeckSnapshot({ formatting: false, current: true })
  const slide = snapshot.slides[0]
  if (!slide) throw new UserFacingError('Select a slide first, in the thumbnails on the left.')
  const label = oneLine(name ?? '') || slideTitle(slide)?.text
  if (!label) throw new UserFacingError('This slide has no title to name the section after. Add a title, then try again.')
  await setSlideTag(slide.id, TAGS.section, label)
  return `Section “${label}” starts at slide ${slide.index + 1}.`
}

export async function renameSection(slideId: string, name: string): Promise<string> {
  const label = oneLine(name)
  if (!label) throw new UserFacingError('Type a name for the section.')
  requirePpt(NEEDS.tags, 'Naming a section')
  await setSlideTag(slideId, TAGS.section, label)
  return `Section renamed “${label}”. Update the trackers to match, and the agenda if there is one.`
}

/** Removes a section mark. Divider slides start a section by being dividers, so they keep theirs. */
export async function removeSectionMark(slideId: string): Promise<string> {
  requirePpt(NEEDS.tags, 'Removing a section')
  await setSlideTag(slideId, TAGS.section, null)
  return 'Section removed. Update the trackers to match, and the agenda if there is one.'
}

export async function insertAgenda(): Promise<string> {
  requirePpt(NEEDS.read, 'Inserting an agenda slide')
  const sections = findSections((await takeDeckSnapshot({ formatting: false })).slides)
  if (!sections.length) throw new NoSectionsError()
  const result = await insertListSlide({ library: ['agenda'], title: 'Agenda', items: sections.map((s) => s.name), numbered: true, layout: 'agenda' })
  return `Agenda added ${result.afterCurrent ? 'after this slide' : 'at the end of the deck'}, listing ${plural(sections.length, 'section')}.`
}

/** Writes a tracker's text, its BwB formatting with the current section in navy bold, and its tag. */
function writeTracker(shape: PowerPoint.Shape, plan: TrackerPlan): void {
  const range = shape.textFrame.textRange
  range.text = plan.text
  range.font.name = TRACKER_STYLE.font
  range.font.size = TRACKER_STYLE.size
  range.font.bold = false
  range.font.color = `#${TRACKER_STYLE.colour}`
  if (plan.current.length) {
    const current = range.getSubstring(plan.current.start, plan.current.length).font
    current.bold = true
    current.color = `#${TRACKER_STYLE.current}`
  }
  shape.tags.add(TAGS.tracker, plan.value)
}

function newTracker(slide: PowerPoint.Slide, plan: TrackerPlan): void {
  const box = slide.shapes.addTextBox(plan.text, { ...TRACKER_BOX })
  box.name = 'BwB section tracker'
  const frame = box.textFrame
  frame.autoSizeSetting = 'AutoSizeNone'
  frame.wordWrap = true
  frame.leftMargin = 0
  frame.rightMargin = 0
  frame.topMargin = 0
  frame.bottomMargin = 0
  frame.verticalAlignment = 'Middle'
  frame.textRange.paragraphFormat.horizontalAlignment = 'Left'
  writeTracker(box, plan)
}

interface TrackerCounts {
  added: number
  updated: number
  removed: number
}

/**
 * Puts an up-to-date tracker on every slide that should have one and removes the rest. Adding needs sections;
 * updating needs trackers, and removes them all once the deck has no sections left.
 */
async function refreshTrackers(mode: 'add' | 'update'): Promise<TrackerCounts> {
  const snapshot = await takeDeckSnapshot({ formatting: false })
  const sections = findSections(snapshot.slides)
  if (mode === 'add' && !sections.length) throw new NoSectionsError()
  if (mode === 'update' && !snapshot.slides.some((s) => trackersOn(s).length)) {
    throw new UserFacingError('The deck has no section trackers to update. Use Add section trackers first.')
  }
  const plans = new Map(planTrackers(snapshot.slides, sections).map((p) => [p.slideId, p]))
  const counts: TrackerCounts = { added: 0, updated: 0, removed: 0 }
  await runPpt(async (context) => {
    for (const slide of snapshot.slides) {
      const target = context.presentation.slides.getItem(slide.id)
      const [keep, ...extra] = trackersOn(slide)
      const plan = plans.get(slide.id)
      const remove = plan ? extra : trackersOn(slide)
      remove.forEach((t) => shapeAt(target, t.path).delete())
      counts.removed += remove.length
      if (!plan) continue
      if (keep) {
        writeTracker(shapeAt(target, keep.path), plan)
        counts.updated += 1
      } else {
        newTracker(target, plan)
        counts.added += 1
      }
    }
    await context.sync()
  })
  return counts
}

function trackerMessage(c: TrackerCounts): string {
  const done = [c.added ? `added to ${plural(c.added, 'slide')}` : '', c.updated ? `updated on ${plural(c.updated, 'slide')}` : '']
    .filter(Boolean)
  const removed = c.removed ? `${plural(c.removed, 'tracker')} removed from slides that do not need one.` : ''
  if (!done.length) return removed ? `No slide needs a section tracker. ${removed}` : 'No slide needs a section tracker yet.'
  return `Section trackers ${done.join(' and ')}.${removed ? ` ${removed}` : ''}`
}

export async function addTrackers(): Promise<string> {
  requirePpt(NEEDS.read, 'Adding section trackers')
  return trackerMessage(await refreshTrackers('add'))
}

export async function updateTrackers(): Promise<string> {
  requirePpt(NEEDS.read, 'Updating section trackers')
  return trackerMessage(await refreshTrackers('update'))
}

export async function removeTrackers(): Promise<string> {
  requirePpt(NEEDS.read, 'Removing section trackers')
  const snapshot = await takeDeckSnapshot({ formatting: false })
  const found = snapshot.slides.flatMap((slide) => trackersOn(slide).map((t) => ({ slide, t })))
  if (!found.length) return 'The deck has no section trackers.'
  await runPpt(async (context) => {
    for (const { slide, t } of found) shapeAt(context.presentation.slides.getItem(slide.id), t.path).delete()
    await context.sync()
  })
  return `Section trackers removed from ${plural(new Set(found.map((f) => f.slide.id)).size, 'slide')}.`
}
