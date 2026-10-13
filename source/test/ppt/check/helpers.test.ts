import { describe, expect, it } from 'vitest'
import { COLOURS } from '../../../src/brand/tokens.ts'
import { SHAPE_COLOURS, TEXT_COLOURS, colourDistance, colourLabel, isLightBackground, isPaletteColour, nearestColour } from '../../../src/ppt/check/colour.ts'
import {
  authoredText, isBwbFurniture, isPicture, isQuoteShape, isSourceShape, placeholderFromName, slideKind, slideTitle, tagValue,
} from '../../../src/ppt/check/shapes.ts'
import { hexColour, mergeSpans, oneLine, points, shorten, splitParagraphs, splitWords, wordCount } from '../../../src/ppt/check/text.ts'
import { foldFindings, runsOf, shapeExcerpt } from '../../../src/ppt/check/util.ts'
import type { DeckFinding } from '../../../src/ppt/check/types.ts'
import { body, font, picture, shape, slide, source, text, title } from './build.ts'

describe('splitting slide text', () => {
  it('splits paragraphs on every break PowerPoint uses, keeping offsets and empty lines', () => {
    const t = 'One\rTwo\nThree\r\nFour\vFive\r'
    const ps = splitParagraphs(t)
    expect(ps.map((p) => p.text)).toEqual(['One', 'Two', 'Three', 'Four', 'Five', ''])
    for (const p of ps) expect(t.slice(p.start, p.start + p.length)).toBe(p.text)
    expect(splitParagraphs('')).toEqual([{ start: 0, length: 0, text: '' }])
  })

  it('splits words with offsets into the whole text', () => {
    const t = 'Alpha  beta\rgamma'
    expect(splitWords(t).map((w) => t.slice(w.start, w.start + w.length))).toEqual(['Alpha', 'beta', 'gamma'])
    expect(splitWords(t, 7, 4)).toEqual([{ start: 7, length: 4 }])
  })

  it('puts a title on one line and counts words', () => {
    expect(oneLine('  Growth\rslowed\v  in 2025 ')).toBe('Growth slowed in 2025')
    expect(oneLine('a\u00A0\u00A0b')).toBe('a b')
    expect(wordCount('  three short words ')).toBe(3)
    expect(wordCount('   ')).toBe(0)
  })

  it('reads hex colours and writes point sizes', () => {
    expect(hexColour('#022338')).toBe('022338')
    expect(hexColour('bed62f')).toBe('BED62F')
    expect(hexColour('')).toBeNull()
    expect(hexColour(null)).toBeNull()
    expect(hexColour('Accent1')).toBeNull()
    expect(points(9)).toBe('9 pt')
    expect(points(9.5)).toBe('9.5 pt')
    expect(points(10.333)).toBe('10.33 pt')
  })

  it('shortens long names at a word boundary', () => {
    expect(shorten('Context', 32)).toBe('Context')
    expect(shorten('Blended finance mobilises private capital for adaptation', 32)).toBe('Blended finance mobilises…')
    expect(shorten('Supercalifragilisticexpialidocious', 10)).toBe('Supercali…')
  })

  it('merges overlapping and touching spans', () => {
    expect(mergeSpans([{ start: 10, length: 5 }, { start: 0, length: 3 }, { start: 3, length: 2 }, { start: 12, length: 10 }, { start: 30, length: 0 }]))
      .toEqual([{ start: 0, length: 5 }, { start: 10, length: 12 }])
  })

  it('merges words with only spaces between them, given the text, but not across lines or other words', () => {
    const t = 'Small  print\rnext line and more'
    const words = [{ start: 0, length: 5 }, { start: 7, length: 5 }, { start: 13, length: 4 }, { start: 27, length: 4 }]
    expect(mergeSpans(words, t)).toEqual([{ start: 0, length: 12 }, { start: 13, length: 4 }, { start: 27, length: 4 }])
    expect(mergeSpans(words)).toEqual(words)
  })
})

describe('the slide palette', () => {
  it('knows the palette', () => {
    expect(isPaletteColour('022338')).toBe(true)
    expect(isPaletteColour('bed62f')).toBe(true)
    expect(isPaletteColour('FF0000')).toBe(false)
    expect(colourLabel('022338')).toBe('Navy')
    expect(colourLabel('D5D5D6')).toBe('Charcoal 20')
    expect(colourLabel('123456')).toBe('#123456')
    expect(isLightBackground('F7F8F5')).toBe(true)
    expect(isLightBackground('022338')).toBe(false)
  })

  it('maps stray colours to the nearest palette colour as the eye sees it', () => {
    expect(colourDistance('022338', '022338')).toBe(0)
    expect(nearestColour('4472C4', TEXT_COLOURS).key).toBe('navy')
    expect(nearestColour('002060', TEXT_COLOURS).key).toBe('navy')
    expect(nearestColour('7F7F7F', TEXT_COLOURS).key).toBe('charcoal')
    expect(nearestColour('00B050', SHAPE_COLOURS).key).toBe('green')
    expect(nearestColour('F2F2F2', SHAPE_COLOURS).key).toBe('offWhite')
    expect(nearestColour('ABCDEF', []).key).toBe('navy')
    expect(nearestColour('ABCDEF', ['black']).hex).toBe(COLOURS.black)
  })
})

describe('what a shape or slide is', () => {
  it('reads placeholder types from PowerPoint’s names when the API cannot say', () => {
    expect(placeholderFromName('Title 1')).toBe('Title')
    expect(placeholderFromName('Subtitle 2')).toBe('Subtitle')
    expect(placeholderFromName('Content Placeholder 2')).toBe('Body')
    expect(placeholderFromName('Text Placeholder 4')).toBe('Body')
    expect(placeholderFromName('Slide Number Placeholder 5')).toBe('SlideNumber')
    expect(placeholderFromName('Footer Placeholder 4')).toBe('Footer')
    expect(placeholderFromName('Date Placeholder 3')).toBe('Date')
    expect(placeholderFromName('Picture Placeholder 6')).toBe('Picture')
    expect(placeholderFromName('Google Shape;55;p13')).toBe('Unsupported')
  })

  it('finds tags whatever the case of the key', () => {
    expect(tagValue({ tags: { BWB_SECTION: 'Context' } }, 'bwb_section')).toBe('Context')
  })

  it('tells sources, furniture, pictures and quotes apart', () => {
    expect(isSourceShape(source('Source: BwB analysis'))).toBe(true)
    expect(isSourceShape(text('Note', { name: 'Footnote 3', top: 200 }))).toBe(true)
    expect(isSourceShape(text('x', { tags: { BWB_ELEMENT: 'source' }, top: 200 }))).toBe(true)
    expect(isSourceShape(text('A message in the content area'))).toBe(false)
    expect(isBwbFurniture(text('DRAFT', { tags: { BWB_STICKER: 'Draft' } }))).toBe(true)
    expect(isBwbFurniture(text('Context | Analysis', { tags: { BWB_TRACKER: '1:Context' } }))).toBe(true)
    expect(isPicture(picture())).toBe(true)
    expect(isPicture(shape({ type: 'Placeholder', placeholder: 'Picture', contains: 'Image' }))).toBe(true)
    expect(isQuoteShape(text('“We need patient capital.”', { tags: { BWB_ELEMENT: 'quote' } }))).toBe(true)
    expect(isQuoteShape(text('We need patient capital.', { font: font({ name: 'Candara', italic: true }) }))).toBe(true)
    expect(isQuoteShape(text('We need patient capital.', { font: font({ italic: true }) }))).toBe(false)
    expect(isQuoteShape(title('An italic title', { font: font({ name: 'Candara', italic: true }) }))).toBe(false)
  })

  it('reads the title from the title placeholder, else the top-most text, skipping furniture and sources', () => {
    expect(slideTitle(slide([body('Body text'), title('Growth slowed\rin 2025')]))).toMatchObject({ text: 'Growth slowed in 2025', fromPlaceholder: true })
    const fallback = slideTitle(slide([
      text('DRAFT', { top: 5, tags: { BWB_STICKER: 'Draft' } }),
      text('Context | Analysis', { top: 8, tags: { BWB_TRACKER: '1:Context' } }),
      source('Source: BwB', { top: 480 }),
      text('Lower text', { top: 200 }),
      text('Right of the top', { top: 40, left: 500 }),
      text('Top text', { top: 40, left: 40 }),
    ]))
    expect(fallback).toMatchObject({ text: 'Top text', fromPlaceholder: false })
    expect(slideTitle(slide([title(''), text('   ')]))).toBeNull()
    expect(slideTitle(slide([picture()]))).toBeNull()
  })

  it('knows title, agenda, divider, summary and closing slides', () => {
    expect(slideKind(slide([], { tags: { BWB_LAYOUT: 'divider' } }))).toBe('divider')
    expect(slideKind(slide([], { tags: { BWB_LAYOUT: 'divider-navy' } }))).toBe('divider')
    expect(slideKind(slide([], { tags: { BWB_LAYOUT: 'title-photo' } }))).toBe('title')
    expect(slideKind(slide([], { tags: { BWB_LAYOUT: 'agenda' } }))).toBe('agenda')
    expect(slideKind(slide([], { tags: { BWB_LAYOUT: 'summary' } }))).toBe('summary')
    expect(slideKind(slide([], { tags: { BWB_LAYOUT: 'closing' } }))).toBe('closing')
    expect(slideKind(slide([], { tags: { BWB_LAYOUT: 'key-figures' } }))).toBe('content')
    expect(slideKind(slide([title('Context')], { layout: 'Section Header' }))).toBe('divider')
    expect(slideKind(slide([title('Context')], { layout: 'BwB Section divider' }))).toBe('divider')
    expect(slideKind(slide([title('Thank you')], { layout: 'BwB Closing' }))).toBe('closing')
    expect(slideKind(slide([title('Deck')], { layout: 'BwB Title' }))).toBe('title')
    expect(slideKind(slide([title('Point')], { layout: 'BwB Title only' }))).toBe('content')
    expect(slideKind(slide([title('Point')], { layout: 'BwB Title and content' }))).toBe('content')
    expect(slideKind(slide([], { tags: { BWB_LAYOUT: 'section-divider' } }))).toBe('divider')
    expect(slideKind(slide([], { tags: { BWB_LAYOUT: 'executive-summary' } }))).toBe('summary')
    expect(slideKind(slide([], { tags: { BWB_LAYOUT: 'disclaimer' } }))).toBe('closing')
    expect(slideKind(slide([], { tags: { BWB_LAYOUT: 'title-and-content' } }))).toBe('content')
    expect(slideKind(slide([title('Executive summary')]))).toBe('summary')
    expect(slideKind(slide([title('Thank you')]))).toBe('closing')
    expect(slideKind(slide([], { layout: 'Title Slide' }))).toBe('title')
    expect(slideKind(slide([title('Deck', { placeholder: 'CenterTitle' })]))).toBe('title')
    expect(slideKind(slide([title('Agenda')]))).toBe('agenda')
    expect(slideKind(slide([title('Growth slowed')]))).toBe('content')
  })

  it('lists the text an author wrote, leaving out furniture and footers', () => {
    const s = slide([
      title('A title'), text('DRAFT', { tags: { BWB_STICKER: 'x' } }),
      shape({ type: 'Placeholder', placeholder: 'SlideNumber', text: '4' }), text(''), text('Body'),
    ])
    expect(authoredText(s).map((x) => x.text)).toEqual(['A title', 'Body'])
  })
})

describe('rule helpers', () => {
  it('uses the pieces of mixed text, else the whole text', () => {
    const whole = text('Hello')
    expect(runsOf(whole)).toEqual([{ start: 0, length: 5, font: whole.font }])
    const mixed = text('Hello world', { font: font({ name: null }), pieces: [{ start: 0, length: 5, font: font({ name: 'Arial' }) }] })
    expect(runsOf(mixed)).toHaveLength(1)
    expect(runsOf(picture())).toEqual([])
    expect(shapeExcerpt(shape({ name: 'Rectangle 4', text: '' }))).toBe('Rectangle 4')
  })

  it('folds many alike findings into one and keeps the rest in place', () => {
    const f = (id: string, key: string): DeckFinding => ({ id, rule: 'r', group: 'brand', severity: 'minor', title: key })
    const list = [f('a', 'x'), f('b', 'y'), f('c', 'x'), f('d', 'x')]
    const folded = foldFindings(list, (x) => x.title, 2, (alike, key) => ({ ...f(`fold-${key}`, key), detail: String(alike.length) }))
    expect(folded.map((x) => x.id)).toEqual(['fold-x', 'b'])
    expect(folded[0]?.detail).toBe('3')
    expect(foldFindings(list, () => null, 0, (alike) => alike[0] ?? f('z', 'z'))).toEqual(list)
  })
})
