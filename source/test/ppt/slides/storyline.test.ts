import { describe, expect, it } from 'vitest'
import { chooseLibrarySlide, listSlidePlan, pickLayout, planListSlide } from '../../../src/ppt/slides/listSlide.ts'
import { outlineOf } from '../../../src/ppt/slides/outline.ts'
import { LONG_TITLE, executiveSummaryLines, storyline, storylineIssues, storylineText } from '../../../src/ppt/slides/storyline.ts'
import { body, deck, shape, slide, text, title } from '../check/build.ts'

const long = 'Blended finance can mobilise far more private capital for adaptation in small island states than it does'

function sampleDeck() {
  return deck([
    slide([title('Adaptation finance', { placeholder: 'CenterTitle' })]),
    slide([title('Agenda')], { tags: { BWB_LAYOUT: 'agenda' } }),
    slide([title('Context')], { tags: { BWB_LAYOUT: 'divider' } }),
    slide([title('Need outstrips flows'), body('Text.')]),
    slide([body('Only body text', { top: 200 })]),
    slide([shape({ type: 'Image', text: null })]),
    slide([title(long)]),
  ], { selected: ['259#1003'] })
}

describe('the storyline', () => {
  it('lists every slide with its number, title, kind and section', () => {
    const entries = storyline(sampleDeck().slides)
    expect(entries.map((e) => [e.number, e.title, e.kind, e.section])).toEqual([
      [1, 'Adaptation finance', 'title', null],
      [2, 'Agenda', 'agenda', null],
      [3, 'Context', 'divider', 'Context'],
      [4, 'Need outstrips flows', 'content', 'Context'],
      [5, 'Only body text', 'content', 'Context'],
      [6, null, 'content', 'Context'],
      [7, long, 'content', 'Context'],
    ])
    expect(entries[4]?.fromPlaceholder).toBe(false)
    expect(entries.map((e) => e.long)).toEqual([false, false, false, false, false, false, true])
    expect(long.length).toBeGreaterThan(LONG_TITLE)
    expect(storylineIssues(entries)).toEqual({ untitled: 1, long: 1 })
  })

  it('copies as plain text with the sections as headings', () => {
    expect(storylineText(storyline(sampleDeck().slides))).toBe([
      '1. Adaptation finance', '2. Agenda', '', 'CONTEXT', '3. Context', '4. Need outstrips flows', '5. Only body text',
      '6. (no title)', `7. ${long}`,
    ].join('\n'))
    expect(storylineText([])).toBe('')
  })

  it('builds the executive summary from the content slides’ titles', () => {
    expect(executiveSummaryLines(storyline(sampleDeck().slides))).toEqual(['Need outstrips flows', 'Only body text', long])
  })
})

describe('the outline the Slides tab shows', () => {
  it('combines the storyline, sections, trackers and current slide from one read', () => {
    const outline = outlineOf(sampleDeck())
    expect(outline.entries).toHaveLength(7)
    expect(outline.sections.map((s) => s.name)).toEqual(['Context'])
    expect(outline.trackers).toEqual({ slides: 0, issues: [] })
    expect(outline.current).toBe('259#1003')
    expect(outlineOf(deck([])).current).toBeNull()
  })
})

describe('list slides', () => {
  it('put the title and list in the placeholders', () => {
    const s = slide([title('', { id: 't' }), body('', { id: 'b' })])
    const plan = listSlidePlan(s)
    expect([plan.title?.id, plan.list?.id]).toEqual(['t', 'b'])
  })

  it('prefer a shape the library tagged for the list, then the largest text shape below the title', () => {
    const tagged = slide([title('', { id: 't' }), body('', { id: 'b' }), text('', { id: 'l', tags: { BWB_ELEMENT: 'agenda-list' } })])
    expect(listSlidePlan(tagged).list?.id).toBe('l')
    const boxes = slide([
      text('', { id: 'small', top: 120, width: 100, height: 50 }), text('', { id: 'big', top: 120, width: 800, height: 300 }),
      text('', { id: 'high', top: 10, width: 900, height: 500 }), shape({ id: 'pic', type: 'Image', text: null, width: 900, height: 400 }),
    ])
    expect([listSlidePlan(boxes).title, listSlidePlan(boxes).list?.id]).toEqual([null, 'big'])
    expect(listSlidePlan(slide([]))).toEqual({ title: null, list: null })
  })

  it('choose a title and content layout, else title only', () => {
    expect(pickLayout([{ id: '1', name: 'Title Slide' }, { id: '2', name: 'Title and Content' }])?.id).toBe('2')
    expect(pickLayout([{ id: '1', name: 'Title Slide' }, { id: '3', name: 'Title Only' }])?.id).toBe('3')
    expect(pickLayout([{ id: '1', name: 'Title Slide' }])).toBeUndefined()
  })
})

/** The BwB library's agenda slide: a title and five rows of number, item, description and rule. */
function agendaSlide() {
  const rows = [1, 2, 3, 4, 5].flatMap((n) => {
    const top = 108 + (n - 1) * 72
    return [
      text(`0${n}`, { id: `n${n}`, name: `Agenda number ${n}`, left: 40, top, width: 58.7, height: 72 }),
      text(`Sample item ${n}`, { id: `i${n}`, name: `Agenda item ${n}`, left: 114.7, top, width: 357.3, height: 72 }),
      text(`Sample description ${n}`, { id: `d${n}`, name: `Agenda description ${n}`, left: 488, top, width: 432, height: 72 }),
      shape({ id: `r${n}`, type: 'Line', name: `Agenda rule ${n}`, text: null, left: 40, top: top + 72, width: 880, height: 0 }),
    ]
  })
  return slide([
    title('Agenda', { id: 't' }), shape({ id: 'top', type: 'Line', name: 'Agenda rule top', text: null }), ...rows,
    shape({ id: 'num', type: 'Placeholder', placeholder: 'SlideNumber', name: 'Slide number', text: '2' }),
  ])
}

/** The BwB library's text slide: title, content placeholder, callout with its bar, and a source line. */
function bulletsSlide() {
  return slide([
    title('Action title: say the so-what', { id: 't' }),
    body('Lead with the point', { id: 'c', left: 40, top: 108, width: 581.3, height: 360 }),
    text('Key takeaway', { id: 'callout', name: 'Callout', left: 637.3, top: 108, width: 282.7, height: 132 }),
    shape({ id: 'bar', type: 'GeometricShape', name: 'Callout bar', text: '', left: 637.3, top: 108, width: 4, height: 132 }),
    text('Source: organisation', { id: 'src', name: 'Source', left: 40, top: 476, width: 700, height: 24 }),
    shape({ id: 'num', type: 'Placeholder', placeholder: 'SlideNumber', name: 'Slide number', text: '5' }),
  ])
}

describe('filling a new list slide', () => {
  it('fills the agenda’s rows, gives the names the whole row, and removes the rows not needed', () => {
    const steps = planListSlide(agendaSlide(), { title: 'Agenda', items: ['Context', 'Options'], numbered: true })
    expect(steps.filter((s) => s.kind === 'text').map((s) => s.kind === 'text' && [s.path[0], s.text])).toEqual([
      ['t', 'Agenda'], ['n1', '01'], ['i1', 'Context'], ['n2', '02'], ['i2', 'Options'],
    ])
    expect(steps.filter((s) => s.kind === 'resize')).toEqual([
      { kind: 'resize', path: ['i1'], left: 114.7, width: 920 - 114.7 },
      { kind: 'resize', path: ['i2'], left: 114.7, width: 920 - 114.7 },
    ])
    const removed = steps.flatMap((s) => (s.kind === 'remove' ? s.path : []))
    expect(removed.sort()).toEqual(['d1', 'd2', 'd3', 'd4', 'd5', 'i3', 'i4', 'i5', 'n3', 'n4', 'n5', 'r3', 'r4', 'r5'].sort())
  })

  it('falls back to one numbered list when there are more sections than rows', () => {
    const items = ['A', 'B', 'C', 'D', 'E', 'F']
    const steps = planListSlide(agendaSlide(), { title: 'Agenda', items, numbered: true })
    expect(steps).toContainEqual({ kind: 'add-list', text: '1. A\n2. B\n3. C\n4. D\n5. E\n6. F', bullets: false })
    const removed = steps.filter((s) => s.kind === 'remove')
    expect(removed).toHaveLength(21)
    expect(removed.some((s) => s.kind === 'remove' && s.path[0] === 'num')).toBe(false)
  })

  it('puts a summary in the content placeholder, widened, and clears the sample callout and source', () => {
    const steps = planListSlide(bulletsSlide(), { title: 'Executive summary', items: ['One', 'Two'], numbered: false })
    expect(steps).toEqual([
      { kind: 'text', path: ['t'], text: 'Executive summary' },
      { kind: 'text', path: ['c'], text: 'One\nTwo', bullets: true },
      { kind: 'resize', path: ['c'], left: 40, width: 880 },
      { kind: 'remove', path: ['callout'] },
      { kind: 'remove', path: ['bar'] },
      { kind: 'remove', path: ['src'] },
    ])
  })

  it('adds a title and list where a plain slide has none', () => {
    expect(planListSlide(slide([]), { title: 'Agenda', items: ['A'], numbered: true })).toEqual([
      { kind: 'add-title', text: 'Agenda' }, { kind: 'add-list', text: '1. A', bullets: false },
    ])
  })

  it('starts from the first library slide the request names', () => {
    const index = { version: 1 as const, file: 'f.pptx', slides: [
      { id: 'bullets', title: 'Text', category: 'Messages', description: '', tags: [], slideId: '260#1', thumb: 't.png' },
    ] }
    expect(chooseLibrarySlide(index, ['content', 'bullets'])?.id).toBe('bullets')
    expect(chooseLibrarySlide(index, ['agenda'])).toBeUndefined()
  })
})
