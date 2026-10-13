import { afterEach, describe, expect, it, vi } from 'vitest'
import { copySlidesToNewDeck } from '../../../src/ppt/slides/copy.ts'
import { PPT_SLIDE_HANDLERS } from '../../../src/ppt/slides/handlers.ts'
import { SlideLibraryError, clearSlideLibraryCache, insertLibrarySlide, loadSlideIndex } from '../../../src/ppt/slides/library.ts'
import { readOutline } from '../../../src/ppt/slides/outline.ts'
import { goToSlide, insertExecutiveSummary, readStoryline } from '../../../src/ppt/slides/storyline.ts'
import type { FakePresentation, FakeSlide } from '../../fakePpt/index.ts'
import {
  createPresentation, installFakePowerPoint, placeholder, selectSlides, shapesTagged, slide, textBox, textOf, uninstallFakePowerPoint,
} from '../../fakePpt/index.ts'
import { LIBRARY_BASE64, serveLibrary, useLibrarySlides } from './fakeLibrary.ts'

afterEach(() => {
  uninstallFakePowerPoint()
  vi.unstubAllGlobals()
  clearSlideLibraryCache()
})

const content = (title: string, id?: string): FakeSlide => slide([placeholder('Title', title), placeholder('Content', 'Text.')], id ? { id } : {})

function storyDeck(): FakePresentation {
  return createPresentation([
    slide([placeholder('CenterTitle', 'Adaptation finance')], { id: 'title', layoutId: 'layout-title' }),
    slide([placeholder('Title', 'Context')], { id: 'd1', layoutId: 'layout-section' }),
    content('Need outstrips flows', 'c1'),
    content('Costs are rising', 'c2'),
    slide([textBox('Only a text box', { top: 200 })], { id: 'c3' }),
  ])
}

describe('inserting a library slide', () => {
  it('goes in after the current slide, tagged with its library id, and is selected', async () => {
    const deck = storyDeck()
    selectSlides(deck, ['c1'])
    useLibrarySlides(deck)
    const fetches = serveLibrary()
    installFakePowerPoint(deck)
    const index = await loadSlideIndex()
    const divider = index.slides.find((s) => s.id === 'section-divider')
    if (!divider) throw new Error('no divider in the index')
    const done = await insertLibrarySlide(divider, 'UseDestinationTheme', index)
    expect(done.afterCurrent).toBe(true)
    const inserted = deck.slides[3]
    expect(done.slideIds).toEqual([inserted?.id])
    expect(inserted?.tags.get('BWB_LAYOUT')).toBe('section-divider')
    expect(deck.selection.slideIds).toEqual([inserted?.id])
    expect(deck.slideInserts).toEqual([{ base64: LIBRARY_BASE64, options: { formatting: 'UseDestinationTheme', sourceSlideIds: ['258#414069683'], targetSlideId: 'c1' } }])
    await insertLibrarySlide(divider, 'KeepSourceFormatting', index)
    expect(fetches.mock.calls.filter(([url]) => String(url).endsWith('.pptx'))).toHaveLength(1)
  })

  it('goes at the end when PowerPoint cannot say which slide is current', async () => {
    const deck = storyDeck()
    useLibrarySlides(deck)
    serveLibrary()
    installFakePowerPoint(deck, { apiVersion: '1.4' })
    const index = await loadSlideIndex()
    const agenda = index.slides[0]
    if (!agenda) throw new Error('no agenda')
    const done = await insertLibrarySlide(agenda, 'KeepSourceFormatting', index)
    expect(done.afterCurrent).toBe(false)
    expect(deck.slides.at(-1)?.tags.get('BWB_LAYOUT')).toBe('agenda')
    expect(deck.slideInserts[0]?.options.targetSlideId).toBe('c3')
  })

  it('says so when PowerPoint inserts nothing', async () => {
    const deck = storyDeck()
    serveLibrary()
    installFakePowerPoint(deck)
    const index = await loadSlideIndex()
    const agenda = index.slides[0]
    if (!agenda) throw new Error('no agenda')
    await expect(insertLibrarySlide(agenda, 'KeepSourceFormatting', index)).rejects.toBeInstanceOf(SlideLibraryError)
  })
})

describe('the storyline in PowerPoint', () => {
  it('reads the titles and goes to a slide', async () => {
    const deck = storyDeck()
    installFakePowerPoint(deck)
    const entries = await readStoryline()
    expect(entries.map((e) => [e.number, e.title, e.kind])).toEqual([
      [1, 'Adaptation finance', 'title'], [2, 'Context', 'divider'], [3, 'Need outstrips flows', 'content'],
      [4, 'Costs are rising', 'content'], [5, 'Only a text box', 'content'],
    ])
    await goToSlide('c2')
    expect(deck.selection.slideIds).toEqual(['c2'])
  })

  it('builds the executive summary from the BwB text slide, clearing its sample callout and source', async () => {
    const deck = storyDeck()
    selectSlides(deck, ['title'])
    useLibrarySlides(deck)
    serveLibrary()
    installFakePowerPoint(deck)
    expect(await insertExecutiveSummary()).toBe('Executive summary added after this slide, with 3 titles as bullets.')
    const summary = deck.slides[1]
    expect(deck.slideInserts[0]?.options).toMatchObject({ sourceSlideIds: ['260#1940440457'] })
    expect(summary?.tags.get('BWB_LAYOUT')).toBe('summary')
    expect(summary?.shapes.map((s) => s.name)).toEqual(['Title', 'Content', 'Slide number'])
    const [title, list] = summary?.shapes ?? []
    expect(title && textOf(title)).toBe('Executive summary')
    expect(list && textOf(list)).toBe('Need outstrips flows\nCosts are rising\nOnly a text box')
    expect(list && [list.left, list.width, list.textFrame?.bulletVisible]).toEqual([40, 880, true])
    // A second summary leaves out the first: it is not a content slide.
    await insertExecutiveSummary()
    expect(textOf(deck.slides[1]?.shapes[1] ?? textBox(''))).toBe('Need outstrips flows\nCosts are rising\nOnly a text box')
  })

  it('builds the summary on a plain slide when the library is not there', async () => {
    const deck = storyDeck()
    selectSlides(deck, ['c3'])
    serveLibrary(null)
    installFakePowerPoint(deck)
    expect(await insertExecutiveSummary()).toBe('Executive summary added after this slide, with 3 titles as bullets.')
    const summary = deck.slides.at(-1)
    expect(summary?.shapes.map((s) => [s.placeholderType, textOf(s), s.textFrame?.bulletVisible])).toEqual([
      ['Title', 'Executive summary', false], ['Content', 'Need outstrips flows\nCosts are rising\nOnly a text box', true],
    ])
  })

  it('needs titled content slides to summarise', async () => {
    serveLibrary(null)
    installFakePowerPoint(createPresentation([slide([placeholder('CenterTitle', 'Deck')], { layoutId: 'layout-title' })]))
    await expect(insertExecutiveSummary()).rejects.toThrow('No content slide has a title yet')
  })
})

describe('the outline the Slides tab reads', () => {
  it('reads the storyline, sections, trackers and current slide in one go', async () => {
    const deck = storyDeck()
    selectSlides(deck, ['c2'])
    deck.slides[2]?.shapes.push(textBox('Context', { tags: { BWB_TRACKER: '1:Context' } }))
    const fake = installFakePowerPoint(deck)
    const outline = await readOutline()
    expect(fake.runs()).toBe(1)
    expect(outline.sections.map((s) => [s.name, s.from])).toEqual([['Context', 'divider']])
    expect(outline.current).toBe('c2')
    expect(outline.trackers.slides).toBe(1)
    expect(outline.trackers.issues.map((i) => [i.slideId, i.problem])).toEqual([['c2', 'missing'], ['c3', 'missing']])
  })
})

describe('copying slides to a new deck', () => {
  it('exports the selected slides and opens them as a new presentation', async () => {
    const deck = storyDeck()
    selectSlides(deck, ['c1', 'c2'])
    installFakePowerPoint(deck)
    expect(await copySlidesToNewDeck()).toBe('2 slides copied to a new presentation. Save it under its own name.')
    expect(deck.slideExports).toEqual([['c1', 'c2']])
    expect(deck.createdPresentations).toEqual([btoa('fake-presentation:c1,c2')])
  })

  it('says plainly when PowerPoint is too old, or nothing is selected', async () => {
    installFakePowerPoint(storyDeck(), { apiVersion: '1.9' })
    await expect(copySlidesToNewDeck()).rejects.toThrow('Copying slides to a new presentation needs a newer PowerPoint')
    uninstallFakePowerPoint()
    const deck = storyDeck()
    selectSlides(deck, [])
    installFakePowerPoint(deck)
    await expect(copySlidesToNewDeck()).rejects.toThrow('Select the slides to copy first')
    expect(deck.createdPresentations).toEqual([])
  })
})

describe('the ribbon’s Agenda menu', () => {
  it('marks a section, adds and updates trackers, inserts the agenda and removes the trackers', async () => {
    const deck = createPresentation([content('Opening', 'a'), content('Context', 'b'), content('More context', 'c'), content('Options', 'd')])
    useLibrarySlides(deck)
    serveLibrary()
    installFakePowerPoint(deck)
    selectSlides(deck, ['b'])
    await PPT_SLIDE_HANDLERS.pptSectionMark()
    selectSlides(deck, ['d'])
    await PPT_SLIDE_HANDLERS.pptSectionMark()
    await PPT_SLIDE_HANDLERS.pptTrackersAdd()
    expect(shapesTagged(deck, 'BWB_TRACKER').map(textOf)).toEqual(['Context | Options', 'Context | Options', 'Context | Options'])
    selectSlides(deck, ['a'])
    await PPT_SLIDE_HANDLERS.pptAgendaInsert()
    expect(deck.slides[1]?.tags.get('BWB_LAYOUT')).toBe('agenda')
    await PPT_SLIDE_HANDLERS.pptTrackersUpdate()
    expect(shapesTagged(deck, 'BWB_TRACKER')).toHaveLength(3)
    await PPT_SLIDE_HANDLERS.pptTrackersRemove()
    expect(shapesTagged(deck, 'BWB_TRACKER')).toEqual([])
  })
})
