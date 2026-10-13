import { afterEach, describe, expect, it, vi } from 'vitest'
import { TRACKER_BOX } from '../../../src/ppt/slides/sections.ts'
import { clearSlideLibraryCache } from '../../../src/ppt/slides/library.ts'
import {
  NoSectionsError, addTrackers, insertAgenda, markSectionHere, removeSectionMark, removeTrackers, renameSection, updateTrackers,
} from '../../../src/ppt/slides/sectionOps.ts'
import type { FakePresentation, FakeShape, FakeSlide } from '../../fakePpt/index.ts'
import {
  createPresentation, installFakePowerPoint, placeholder, selectSlides, shapesTagged, slide, textBox, textFrame, textOf, uninstallFakePowerPoint,
} from '../../fakePpt/index.ts'
import { serveLibrary, useLibrarySlides } from './fakeLibrary.ts'

afterEach(() => {
  uninstallFakePowerPoint()
  vi.unstubAllGlobals()
  clearSlideLibraryCache()
})

const content = (title: string, id?: string): FakeSlide => slide([placeholder('Title', title), placeholder('Content', 'Text.')], id ? { id } : {})
const divider = (title: string, id?: string): FakeSlide => slide([placeholder('Title', title)], { layoutId: 'layout-section', ...(id ? { id } : {}) })

/** Title slide, two divider-led sections with two and one content slides, and a closing slide. */
function sectionedDeck(): FakePresentation {
  return createPresentation([
    slide([placeholder('CenterTitle', 'Adaptation finance')], { id: 'title', layoutId: 'layout-title' }),
    divider('Context', 'd1'),
    content('Need outstrips flows', 'c1'),
    content('Costs are rising', 'c2'),
    divider('Options', 'd2'),
    content('Guarantees work', 'c3'),
    slide([placeholder('Title', 'Thank you')], { id: 'end' }),
  ])
}

const trackerOn = (s: FakeSlide | undefined): FakeShape[] => (s ? shapesTagged(s, 'BWB_TRACKER') : [])
const slideById = (deck: FakePresentation, id: string): FakeSlide | undefined => deck.slides.find((s) => s.id === id)

describe('marking sections', () => {
  it('starts a section at the current slide, named after its title or as given', async () => {
    const deck = sectionedDeck()
    selectSlides(deck, ['c2'])
    installFakePowerPoint(deck)
    expect(await markSectionHere()).toBe('Section “Costs are rising” starts at slide 4.')
    expect(slideById(deck, 'c2')?.tags.get('BWB_SECTION')).toBe('Costs are rising')
    expect(await markSectionHere('  Costs\rand risks ')).toBe('Section “Costs and risks” starts at slide 4.')
    expect(slideById(deck, 'c2')?.tags.get('BWB_SECTION')).toBe('Costs and risks')
  })

  it('renames and removes a section mark', async () => {
    const deck = sectionedDeck()
    installFakePowerPoint(deck)
    await renameSection('d1', 'Where we are')
    expect(slideById(deck, 'd1')?.tags.get('BWB_SECTION')).toBe('Where we are')
    await removeSectionMark('d1')
    expect(slideById(deck, 'd1')?.tags.has('BWB_SECTION')).toBe(false)
    await expect(renameSection('d1', '  ')).rejects.toThrow('Type a name for the section.')
  })

  it('needs a title to name the section after, and PowerPoint that knows the current slide', async () => {
    const deck = createPresentation([slide([textBox('')], { id: 'blank' })])
    installFakePowerPoint(deck)
    await expect(markSectionHere()).rejects.toThrow('This slide has no title to name the section after.')
    uninstallFakePowerPoint()
    installFakePowerPoint(sectionedDeck(), { apiVersion: '1.4' })
    await expect(markSectionHere()).rejects.toThrow('Starting a section at the current slide needs a newer PowerPoint')
  })
})

describe('section trackers', () => {
  it('go on every content slide from the first section on, current section in navy bold, the rest charcoal 9 pt', async () => {
    const deck = sectionedDeck()
    installFakePowerPoint(deck)
    expect(await addTrackers()).toBe('Section trackers added to 3 slides.')
    expect(deck.slides.map((s) => trackerOn(s).length)).toEqual([0, 0, 1, 1, 0, 1, 0])
    const [tracker] = trackerOn(slideById(deck, 'c3'))
    expect(tracker && textOf(tracker)).toBe('Context | Options')
    expect(tracker?.tags.get('BWB_TRACKER')).toBe('2:Options')
    expect(tracker?.textFrame?.runs.map((r) => [r.text, r.font.name, r.font.size, r.font.bold, r.font.color])).toEqual([
      ['Context | ', 'Aptos', 9, false, '#4A4A4C'],
      ['Options', 'Aptos', 9, true, '#022338'],
    ])
    expect(tracker && [tracker.left, tracker.top, tracker.width, tracker.height]).toEqual([TRACKER_BOX.left, TRACKER_BOX.top, TRACKER_BOX.width, TRACKER_BOX.height])
    expect(tracker?.textFrame).toMatchObject({ autoSizeSetting: 'AutoSizeNone', leftMargin: 0, topMargin: 0, verticalAlignment: 'Middle', horizontalAlignment: 'Left' })
  })

  it('update in place, add missing ones and remove those that no longer belong', async () => {
    const deck = sectionedDeck()
    installFakePowerPoint(deck)
    await addTrackers()
    const kept = trackerOn(slideById(deck, 'c1'))[0]
    const renamed = slideById(deck, 'd2')?.shapes[0]
    if (renamed) renamed.textFrame = textFrame('Choices')
    deck.slides.splice(4, 0, content('A new slide', 'new'))
    slideById(deck, 'title')?.shapes.push(textBox('stray', { tags: { BWB_TRACKER: '1:Context' } }))
    expect(await updateTrackers()).toBe('Section trackers added to 1 slide and updated on 3 slides. 1 tracker removed from slides that do not need one.')
    expect(trackerOn(slideById(deck, 'c1'))[0]).toBe(kept)
    expect(kept && textOf(kept)).toBe('Context | Choices')
    expect(trackerOn(slideById(deck, 'new')).map(textOf)).toEqual(['Context | Choices'])
    expect(trackerOn(slideById(deck, 'title'))).toEqual([])
  })

  it('are removed by Update once the deck has no sections, and Update needs trackers to start from', async () => {
    const deck = sectionedDeck()
    installFakePowerPoint(deck)
    await expect(updateTrackers()).rejects.toThrow('The deck has no section trackers to update.')
    await addTrackers()
    for (const s of deck.slides) s.layoutId = 'layout-title-content'
    expect(await updateTrackers()).toBe('No slide needs a section tracker. 3 trackers removed from slides that do not need one.')
    expect(shapesTagged(deck, 'BWB_TRACKER')).toEqual([])
  })

  it('come off every slide in one go', async () => {
    const deck = sectionedDeck()
    installFakePowerPoint(deck)
    expect(await removeTrackers()).toBe('The deck has no section trackers.')
    await addTrackers()
    expect(await removeTrackers()).toBe('Section trackers removed from 3 slides.')
    expect(shapesTagged(deck, 'BWB_TRACKER')).toEqual([])
  })

  it('need sections first', async () => {
    installFakePowerPoint(createPresentation([content('One'), content('Two')]))
    await expect(addTrackers()).rejects.toBeInstanceOf(NoSectionsError)
  })
})

describe('the agenda slide', () => {
  it('fills the BwB library agenda’s rows after the current slide', async () => {
    const deck = sectionedDeck()
    selectSlides(deck, ['title'])
    useLibrarySlides(deck)
    serveLibrary()
    installFakePowerPoint(deck)
    expect(await insertAgenda()).toBe('Agenda added after this slide, listing 2 sections.')
    const agenda = deck.slides[1]
    expect(deck.slideInserts[0]?.options).toMatchObject({ formatting: 'KeepSourceFormatting', sourceSlideIds: ['257#136333637'], targetSlideId: 'title' })
    expect(agenda?.tags.get('BWB_LAYOUT')).toBe('agenda')
    expect(deck.selection.slideIds).toEqual([agenda?.id])
    const named = (name: string) => agenda?.shapes.find((s) => s.name === name)
    expect(['Agenda number 1', 'Agenda item 1', 'Agenda number 2', 'Agenda item 2'].map((n) => { const s = named(n); return s ? textOf(s) : null }))
      .toEqual(['01', 'Context', '02', 'Options'])
    expect(named('Agenda item 1')?.width).toBeCloseTo(920 - 114.7)
    expect(agenda?.shapes.map((s) => s.name)).toEqual([
      'Title', 'Agenda rule top', 'Agenda number 1', 'Agenda item 1', 'Agenda rule 1', 'Agenda number 2', 'Agenda item 2', 'Agenda rule 2', 'Slide number',
    ])
  })

  it('builds one on a plain slide when the library is not there, next to the current slide', async () => {
    const deck = sectionedDeck()
    selectSlides(deck, ['title'])
    serveLibrary(null)
    installFakePowerPoint(deck)
    expect(await insertAgenda()).toBe('Agenda added after this slide, listing 2 sections.')
    const agenda = deck.slides[1]
    expect(agenda?.layoutId).toBe('layout-title-content')
    expect(agenda?.shapes.map((s) => [s.placeholderType, textOf(s), s.textFrame?.bulletVisible])).toEqual([
      ['Title', 'Agenda', false], ['Content', '1. Context\n2. Options', false],
    ])
    expect(agenda?.tags.get('BWB_LAYOUT')).toBe('agenda')
  })

  it('goes at the end of the deck when PowerPoint cannot move slides', async () => {
    const deck = sectionedDeck()
    selectSlides(deck, ['title'])
    serveLibrary(null)
    installFakePowerPoint(deck, { apiVersion: '1.5' })
    expect(await insertAgenda()).toBe('Agenda added at the end of the deck, listing 2 sections.')
    expect(deck.slides.at(-1)?.tags.get('BWB_LAYOUT')).toBe('agenda')
  })

  it('needs sections', async () => {
    installFakePowerPoint(createPresentation([content('One')]))
    await expect(insertAgenda()).rejects.toBeInstanceOf(NoSectionsError)
  })
})
