import { describe, expect, it } from 'vitest'
import {
  TRACKER_SEPARATOR, agendaLines, findSections, planTrackers, sectionOf, trackerIssues, trackerText, trackerValue,
} from '../../../src/ppt/slides/sections.ts'
import { body, deck, slide, text, title } from '../check/build.ts'

const divider = (name: string) => slide([title(name)], { tags: { BWB_LAYOUT: 'divider' } })
const content = (name: string, over: Parameters<typeof slide>[1] = {}) => slide([title(name), body('Text.')], over)
const tracker = (t: string, value: string) => text(t, { tags: { BWB_TRACKER: value }, top: 8 })

/** Title, agenda, two divider-led sections and a closing slide. */
function dividerDeck() {
  return deck([
    slide([title('Adaptation finance', { placeholder: 'CenterTitle' })]),
    slide([title('Agenda')], { tags: { BWB_LAYOUT: 'agenda' } }),
    divider('Context'),
    content('Need outstrips flows'),
    content('Costs are rising'),
    divider('Options'),
    content('Guarantees work'),
    slide([title('Thank you')], { tags: { BWB_LAYOUT: 'closing' } }),
  ])
}

describe('sections', () => {
  it('start at divider slides, named after their titles', () => {
    expect(findSections(dividerDeck().slides)).toEqual([
      { name: 'Context', slideId: '258#1002', index: 2, from: 'divider' },
      { name: 'Options', slideId: '261#1005', index: 5, from: 'divider' },
    ])
  })

  it('start at slides marked with a section name, which wins over a divider’s title', () => {
    const snap = deck([
      content('Opening'),
      content('Need outstrips flows', { tags: { BWB_SECTION: '  The\rproblem ' } }),
      slide([title('Ignored title')], { tags: { BWB_LAYOUT: 'divider', BWB_SECTION: 'Options' } }),
      slide([], { tags: { BWB_LAYOUT: 'divider' } }),
      content('Blank tag', { tags: { BWB_SECTION: '' } }),
    ])
    expect(findSections(snap.slides).map((s) => [s.name, s.index, s.from])).toEqual([
      ['The problem', 1, 'tag'], ['Options', 2, 'tag'], ['Section 3', 3, 'divider'],
    ])
  })

  it('number the agenda and tell which section a slide is in', () => {
    const sections = findSections(dividerDeck().slides)
    expect(agendaLines(sections)).toEqual(['1. Context', '2. Options'])
    expect([0, 2, 4, 5, 7].map((i) => sectionOf(sections, i))).toEqual([-1, 0, 0, 1, 1])
  })
})

describe('trackers', () => {
  it('list the sections and mark where the current one is', () => {
    const sections = findSections(dividerDeck().slides)
    const t = trackerText(sections, 1)
    expect(t.text).toBe(`Context${TRACKER_SEPARATOR}Options`)
    expect(t.text.slice(t.current.start, t.current.start + t.current.length)).toBe('Options')
    expect(trackerValue(sections, 1)).toBe('2:Options')
  })

  it('shorten long section names', () => {
    const snap = deck([content('A long action title', { tags: { BWB_SECTION: 'Blended finance mobilises private capital for adaptation' } })])
    const t = trackerText(findSections(snap.slides), 0)
    expect(t.text).toBe('Blended finance mobilises…')
    expect(t.current).toEqual({ start: 0, length: t.text.length })
  })

  it('go on content slides from the first section on, not on title, agenda, divider or closing slides', () => {
    const plans = planTrackers(dividerDeck().slides)
    expect(plans.map((p) => [p.index, p.section, p.value])).toEqual([[3, 0, '1:Context'], [4, 0, '1:Context'], [6, 1, '2:Options']])
    expect(planTrackers(deck([content('No sections')]).slides)).toEqual([])
  })

  it('include a marked section start that is itself a content slide', () => {
    const snap = deck([content('Opening'), content('Start', { tags: { BWB_SECTION: 'Context' } }), content('Next')])
    expect(planTrackers(snap.slides).map((p) => p.index)).toEqual([1, 2])
  })

  it('are up to date only when every planned slide has one tracker with the right text and section', () => {
    const base = dividerDeck()
    expect(trackerIssues(base.slides)).toEqual([])

    const withTrackers = (fn: (index: number) => ReturnType<typeof tracker>[]) =>
      base.slides.map((s) => ({ ...s, shapes: [...s.shapes, ...fn(s.index)] }))
    const good = (index: number) => {
      const plan = planTrackers(base.slides).find((p) => p.index === index)
      return plan ? [tracker(plan.text, plan.value)] : []
    }
    expect(trackerIssues(withTrackers(good))).toEqual([])

    const issues = trackerIssues(withTrackers((i) => {
      if (i === 3) return []
      if (i === 4) return [tracker('Context | Options', '2:Options')]
      if (i === 6) return [...good(6), ...good(6)]
      if (i === 2) return [tracker('Context | Options', '1:Context')]
      return good(i)
    }))
    expect(issues.map((x) => [x.index, x.problem])).toEqual([[2, 'extra'], [3, 'missing'], [4, 'outdated'], [6, 'outdated']])
  })
})
