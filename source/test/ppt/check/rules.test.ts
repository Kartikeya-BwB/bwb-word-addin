import { describe, expect, it } from 'vitest'
import { TITLE_AREA } from '../../../src/brand/slideGrid.ts'
import { brandFindings } from '../../../src/ppt/check/rulesBrand.ts'
import { layoutFindings, usualTitleBox } from '../../../src/ppt/check/rulesLayout.ts'
import { sendFindings } from '../../../src/ppt/check/rulesSend.ts'
import { slideFindings } from '../../../src/ppt/check/rulesSlides.ts'
import { wordingFindings } from '../../../src/ppt/check/rulesWording.ts'
import type { DeckFinding, DeckSnapshot } from '../../../src/ppt/check/types.ts'
import { body, deck, font, picture, shape, slide, source, text, title } from './build.ts'

const rules = (findings: DeckFinding[]): string[] => findings.map((f) => f.rule)
const one = (findings: DeckFinding[], rule: string): DeckFinding | undefined => findings.find((f) => f.rule === rule)
const onOneSlide = (...shapes: Parameters<typeof slide>[0]): DeckSnapshot => deck([slide(shapes)])

describe('brand: fonts', () => {
  it('flags fonts other than Candara and Aptos, offering Candara for titles and Aptos for the rest', () => {
    const snap = onOneSlide(
      title('Growth slowed', { font: font({ name: 'Calibri Light', size: 24 }) }),
      text('Body in Arial', { font: font({ name: 'Arial' }) }),
    )
    const found = brandFindings(snap).filter((f) => f.rule === 'brand/font')
    expect(found.map((f) => f.title)).toEqual(['Text in Calibri Light', 'Text in Arial'])
    expect(found[0]?.fix).toMatchObject({ kind: 'font', font: 'Candara', spans: [{ start: 0, length: 13 }], expectedText: 'Growth slowed' })
    expect(found[1]?.fix).toMatchObject({ kind: 'font', font: 'Aptos' })
    expect(found[1]?.fixLabel).toBe('Change to Aptos')
    expect(found[1]?.slide).toBe(1)
  })

  it('allows the BwB fonts, symbol fonts, theme references and mixed text it could not resolve', () => {
    const snap = onOneSlide(
      text('a', { font: font({ name: 'Aptos' }) }), text('b', { font: font({ name: 'Candara' }) }),
      text('c', { font: font({ name: 'Aptos Display' }) }), text('d', { font: font({ name: 'Wingdings' }) }),
      text('e', { font: font({ name: 'Symbol' }) }), text('f', { font: font({ name: '+mn-lt' }) }), text('g', { font: font({ name: null }) }),
    )
    expect(rules(brandFindings(snap))).not.toContain('brand/font')
  })

  it('fixes only the pieces of mixed text that are off-brand', () => {
    const snap = onOneSlide(text('Hello world', {
      font: font({ name: null }),
      pieces: [{ start: 0, length: 5, font: font({ name: 'Aptos' }) }, { start: 6, length: 5, font: font({ name: 'Arial' }) }],
    }))
    expect(one(brandFindings(snap), 'brand/font')?.fix).toMatchObject({ spans: [{ start: 6, length: 5 }] })
  })

  it('folds a deck set in another font into one finding that fixes every shape', () => {
    const slides = Array.from({ length: 8 }, (_, i) => slide([
      title(`Title ${i}`, { font: font({ name: 'Calibri', size: 24 }) }), text(`Body ${i}`, { font: font({ name: 'Calibri' }) }),
    ]))
    const found = brandFindings(deck(slides)).filter((f) => f.rule.includes('font'))
    expect(found).toHaveLength(1)
    expect(found[0]).toMatchObject({ rule: 'brand/deck-font', title: 'Calibri in 16 shapes on 8 slides', slide: 1 })
    expect(found[0]?.fix?.kind).toBe('batch')
    const fixes = found[0]?.fix?.kind === 'batch' ? found[0].fix.fixes : []
    expect(fixes).toHaveLength(16)
    expect(fixes.filter((f) => f.kind === 'font' && f.font === 'Candara')).toHaveLength(8)
  })

  it('checks nothing when the snapshot holds no formatting', () => {
    const snap = deck([slide([text('Arial', { font: font({ name: 'Arial', size: 6, colour: 'FF0000' }) })])], {
      can: { formatting: false, placeholders: true, groups: true, altText: false, pageSize: true, selection: true },
    })
    expect(brandFindings(snap)).toEqual([])
  })
})

describe('brand: colours', () => {
  it('flags text colours outside the palette and offers the nearest text colour', () => {
    const f = one(brandFindings(onOneSlide(text('Alert', { font: font({ colour: 'FF0000' }) }))), 'brand/off-palette')
    expect(f?.title).toBe('Text colour #FF0000 is not in the BwB palette')
    expect(f?.fix).toMatchObject({ kind: 'text-colour', colour: '4A4A4C', spans: [{ start: 0, length: 5 }] })
    expect(f?.fixLabel).toBe('Change to Charcoal')
  })

  it('flags fills and outlines outside the palette, never offering lime', () => {
    const found = brandFindings(onOneSlide(
      shape({ type: 'GeometricShape', text: '', fill: { type: 'Solid', colour: 'FFFF00' }, line: { visible: true, colour: '5B9BD5' } }),
      shape({ type: 'GeometricShape', text: '', fill: { type: 'Solid', colour: '022338' }, line: { visible: false, colour: 'FF00FF' } }),
    ))
    const fill = found.find((f) => f.fix?.kind === 'fill-colour')
    const line = found.find((f) => f.fix?.kind === 'line-colour')
    expect(fill?.title).toBe('Fill #FFFF00 is not in the BwB palette')
    expect(fill?.fix).not.toMatchObject({ colour: 'BED62F' })
    expect(line?.title).toBe('Outline #5B9BD5 is not in the BwB palette')
    expect(found).toHaveLength(2)
  })

  it('leaves red, amber and green status dots and pictures alone', () => {
    const found = brandFindings(onOneSlide(
      shape({ type: 'GeometricShape', text: '', tags: { BWB_RAG: 'amber' }, fill: { type: 'Solid', colour: 'FFC000' }, line: { visible: true, colour: 'FFC000' } }),
      picture({ fill: { type: 'Solid', colour: 'FF0000' } }),
      shape({ type: 'GeometricShape', text: '', fill: { type: 'Gradient', colour: 'FF0000' } }),
    ))
    expect(found).toEqual([])
  })

  it('flags lime text on a light fill, but not on navy or with no fill', () => {
    const lime = font({ colour: 'BED62F' })
    const found = brandFindings(onOneSlide(
      text('On white', { font: lime, fill: { type: 'Solid', colour: 'FFFFFF' } }),
      text('On navy', { font: lime, fill: { type: 'Solid', colour: '022338' } }),
      text('No fill', { font: lime }),
    ))
    expect(rules(found)).toEqual(['brand/lime-on-light'])
    expect(found[0]?.fix).toMatchObject({ kind: 'text-colour', colour: '022338' })
  })

  it('folds one stray colour used across the deck', () => {
    const slides = Array.from({ length: 16 }, (_, i) => slide([text(`Point ${i}`, { font: font({ colour: '0070C0' }) })]))
    const found = brandFindings(deck(slides))
    expect(found).toHaveLength(1)
    expect(found[0]).toMatchObject({ rule: 'brand/deck-colour', title: 'Text colour #0070C0 in 16 shapes on 16 slides', fixLabel: 'Change them all to Navy' })
  })
})

describe('brand: small text', () => {
  it('flags text below 10 pt, and sources below 8 pt', () => {
    const found = brandFindings(onOneSlide(
      text('Small body', { font: font({ size: 9 }) }),
      source('Source: fine at 8'),
      source('Source: too small', { font: font({ size: 7 }) }),
      text('Context | Analysis', { tags: { BWB_TRACKER: '1:Context' }, font: font({ size: 9 }) }),
      shape({ type: 'Placeholder', placeholder: 'SlideNumber', text: '4', font: font({ size: 8 }), top: 510 }),
    )).filter((f) => f.rule === 'brand/small-text')
    expect(found.map((f) => [f.title, f.fixLabel])).toEqual([['Text at 9 pt', 'Set to 10 pt'], ['Text at 7 pt', 'Set to 8 pt']])
    expect(found[0]?.fix).toMatchObject({ kind: 'font-size', size: 10, spans: [{ start: 0, length: 10 }] })
  })

  it('looks at the pieces of mixed text and ignores blank ones', () => {
    const found = brandFindings(onOneSlide(text('Big small', {
      font: font({ size: null }),
      pieces: [{ start: 0, length: 3, font: font({ size: 14 }) }, { start: 4, length: 5, font: font({ size: 9.5 }) }],
    }), text('   ', { font: font({ size: 6 }) })))
    expect(found.map((f) => f.title)).toEqual(['Text at 9.5 pt'])
  })
})

describe('layout', () => {
  it('flags shapes parked off the slide, with a fix that deletes them', () => {
    const f = one(layoutFindings(onOneSlide(text('Parked', { left: 1000 }))), 'layout/off-slide')
    expect(f?.fix).toMatchObject({ kind: 'delete-shape', from: { left: 1000, top: 108, width: 400, height: 100 } })
  })

  it('moves a shape that runs off the edge back onto the slide, and only suggests it for photos', () => {
    const found = layoutFindings(onOneSlide(
      text('Over the edge', { left: 900, width: 100 }),
      picture({ left: -20, top: 0, width: 500, height: 540 }),
      text('Too big to fit', { left: -10, width: 1000 }),
      text('Rounding', { left: 560.3, width: 400 }),
      text('In a group', { left: 950, level: 1, path: ['g', 'c'] }),
    )).filter((f) => f.rule === 'layout/off-edge')
    expect(found).toHaveLength(3)
    expect(found[0]?.fix).toMatchObject({ kind: 'move-onto-slide', to: { left: 860, top: 108 } })
    expect(found[1]).toMatchObject({ advisory: true })
    expect(found[2]?.fix).toBeUndefined()
  })

  it('judges the edges against the deck’s own slide size', () => {
    const snap = deck([slide([text('Wide', { left: 700, width: 100 })])], { width: 720 })
    expect(rules(layoutFindings(snap))).toContain('layout/off-edge')
  })

  it('flags empty placeholders but leaves titles, footers and unread ones alone', () => {
    const found = layoutFindings(onOneSlide(
      body(''), title(''), shape({ type: 'Placeholder', placeholder: 'Footer', text: '' }),
      body('', { text: null }), shape({ type: 'Placeholder', placeholder: 'Unsupported', text: '' }),
      shape({ type: 'Placeholder', placeholder: 'Picture', text: null, contains: null }),
      shape({ type: 'Placeholder', placeholder: 'Picture', text: null, contains: 'Image' }),
      body('', { contains: 'Table' }),
    )).filter((f) => f.rule === 'layout/empty-placeholder')
    expect(found.map((f) => f.fix)).toEqual([
      expect.objectContaining({ kind: 'delete-empty', textual: true }),
      expect.objectContaining({ kind: 'delete-empty', textual: false }),
    ])
  })

  it('needs placeholder contents from PowerPoint before calling a picture placeholder empty', () => {
    const snap = deck([slide([shape({ type: 'Placeholder', placeholder: 'Picture', text: null })])], {
      can: { formatting: true, placeholders: false, groups: false, altText: false, pageSize: false, selection: true },
    })
    expect(layoutFindings(snap)).toEqual([])
  })

  it('flags empty text boxes with no fill or outline', () => {
    const found = layoutFindings(onOneSlide(
      text(''), text(' ', { fill: { type: 'Solid', colour: '022338' } }), text('', { line: { visible: true, colour: '022338' } }),
      text('', { fill: null }), text('', { tags: { BWB_NOTE: 'x' } }), text('Words'),
    )).filter((f) => f.rule === 'layout/empty-text-box')
    expect(found).toHaveLength(1)
    expect(found[0]?.fix).toMatchObject({ kind: 'delete-empty', textual: true })
  })

  it('finds titles that jump and lines them up with the usual position', () => {
    const slides = [
      slide([title('Deck', { placeholder: 'CenterTitle', top: 200 })]),
      slide([title('One')]), slide([title('Two')]), slide([title('Three', { top: 29.5 })]),
      slide([title('Jumps', { top: 40, height: 80 })]),
      slide([title('Divider', { top: 300 })], { tags: { BWB_LAYOUT: 'divider' } }),
    ]
    const found = layoutFindings(deck(slides)).filter((f) => f.rule === 'layout/jumping-title')
    expect(found).toHaveLength(1)
    expect(found[0]?.slide).toBe(5)
    expect(found[0]?.detail).toContain('12 pt lower and a different size')
    expect(found[0]?.fix).toMatchObject({ kind: 'align-title', from: { top: 40, height: 80 }, to: { ...TITLE_AREA } })
  })

  it('needs at least three titles and a clear usual position', () => {
    expect(layoutFindings(deck([slide([title('A')]), slide([title('B', { top: 60 })])]))).toEqual([])
    expect(usualTitleBox([{ ...TITLE_AREA }, { ...TITLE_AREA, top: 60 }, { ...TITLE_AREA, top: 90 }])).toBeNull()
    const tie = usualTitleBox([{ ...TITLE_AREA, top: 60 }, { ...TITLE_AREA, top: 60 }, { ...TITLE_AREA }, { ...TITLE_AREA }])
    expect(tie?.box.top).toBe(TITLE_AREA.top)
  })
})

describe('slides', () => {
  it('needs a title on every slide', () => {
    const f = one(slideFindings(onOneSlide(picture())), 'slides/no-title')
    expect(f).toMatchObject({ severity: 'major', location: { kind: 'slide', slide: 1 } })
  })

  it('wants the title in the title placeholder', () => {
    expect(rules(slideFindings(onOneSlide(title(''), text('Typed in a box', { top: 30 }))))).toEqual(['slides/empty-title'])
    const boxed = one(slideFindings(onOneSlide(text('Typed in a box', { top: 30 }))), 'slides/title-not-placeholder')
    expect(boxed).toMatchObject({ advisory: true, excerpt: 'Typed in a box' })
  })

  it('flags titles over two lines', () => {
    const long = 'Blended finance can mobilise far more private capital for adaptation in small island states than it does today'
    const f = one(slideFindings(onOneSlide(title(long))), 'slides/long-title')
    expect(f?.title).toBe(`Title over two lines (${long.length} characters)`)
    expect(rules(slideFindings(onOneSlide(title('A short action title'))))).toEqual([])
  })

  it('suggests trimming slides over 120 words, counting only what the author wrote', () => {
    const words = (n: number): string => Array.from({ length: n }, () => 'word').join(' ')
    expect(one(slideFindings(onOneSlide(title('T'), body(words(120)))), 'slides/too-much-text')).toMatchObject({ advisory: true, title: '121 words on the slide' })
    expect(rules(slideFindings(onOneSlide(title('T'), body(words(100)), text(words(50), { tags: { BWB_NOTE: 'n' } }))))).toEqual([])
  })

  it('asks for alt text on graphics such as icons too', () => {
    const found = slideFindings(onOneSlide(
      title('T'), shape({ type: 'Graphic', name: 'Graphic 4', text: null, alt: { title: '', description: '', decorative: false } }),
      shape({ type: 'Graphic', text: null, alt: { title: '', description: 'A green tick', decorative: false } }),
      shape({ type: 'Graphic', text: null, alt: { title: '', description: '', decorative: true } }),
    )).filter((f) => f.rule === 'slides/alt-text')
    expect(found.map((f) => [f.title, f.excerpt])).toEqual([['Graphic without alt text', 'Graphic 4']])
    expect(found[0]?.detail).toContain('icon or graphic')
  })

  it('asks for alt text on pictures unless they are decorative or PowerPoint cannot say', () => {
    const found = slideFindings(onOneSlide(
      title('T'), picture(), picture({ alt: { title: '', description: '', decorative: true } }),
      picture({ alt: { title: '', description: 'Solar panels on a roof', decorative: false } }), picture({ alt: null }),
      shape({ type: 'Placeholder', placeholder: 'Picture', contains: 'Image', text: null, alt: { title: '', description: '', decorative: false } }),
    )).filter((f) => f.rule === 'slides/alt-text')
    expect(found).toHaveLength(2)
  })
})

describe('before you send', () => {
  it('collects review notes into one major finding that removes them all', () => {
    const snap = deck([
      slide([title('One'), text('Check this figure', { tags: { BWB_NOTE: '1' } })]),
      slide([title('Two'), text('Source?', { tags: { BWB_NOTE: '1' } }), text('And this', { tags: { BWB_NOTE: '1' } })]),
    ])
    const f = one(sendFindings(snap), 'send/review-notes')
    expect(f).toMatchObject({ severity: 'major', title: '3 review notes left on 2 slides', excerpt: 'Check this figure', fix: { kind: 'delete-tagged', tag: 'BWB_NOTE' } })
  })

  it('lists each sticker label once, as a suggestion', () => {
    const sticker = (label: string) => text(label, { tags: { BWB_STICKER: label } })
    const snap = deck([slide([title('A'), sticker('DRAFT')]), slide([title('B'), sticker('Draft')]), slide([title('C'), sticker('CONFIDENTIAL')])])
    const found = sendFindings(snap).filter((f) => f.rule === 'send/sticker')
    expect(found.map((f) => f.title)).toEqual(['“DRAFT” sticker on 2 slides', '“CONFIDENTIAL” sticker on 1 slide'])
    expect(found[0]).toMatchObject({ advisory: true, fix: { kind: 'delete-tagged', tag: 'BWB_STICKER', text: 'DRAFT' } })
  })

  it('flags trackers that no longer match the sections', () => {
    const tracker = (t: string, value: string) => text(t, { tags: { BWB_TRACKER: value }, top: 8 })
    const fresh = deck([
      slide([title('Deck', { placeholder: 'CenterTitle' })]),
      slide([title('Context')], { tags: { BWB_LAYOUT: 'divider' } }),
      slide([title('Point'), tracker('Context | Analysis', '1:Context')]),
      slide([title('Analysis')], { tags: { BWB_LAYOUT: 'divider' } }),
      slide([title('Point'), tracker('Context | Analysis', '2:Analysis')]),
    ])
    expect(sendFindings(fresh)).toEqual([])
    const renamed = { ...fresh, slides: fresh.slides.map((s) => (s.index === 3 ? { ...s, shapes: [title('Findings')] } : s)) }
    const f = one(sendFindings(renamed), 'send/trackers')
    expect(f).toMatchObject({ title: 'Section trackers out of date on 2 slides', slide: 3, fix: { kind: 'update-trackers' } })
  })
})

describe('wording', () => {
  it('runs the house-style rules on each paragraph, with offsets into the whole shape text', () => {
    const t = 'First point\rThe organization will analyze it'
    const found = wordingFindings(onOneSlide(body(t)))
    expect(found.map((f) => f.title)).toEqual(['US spelling', 'US spelling'])
    const fixes = found.map((f) => f.fix)
    for (const fix of fixes) {
      if (fix?.kind !== 'replace') throw new Error('expected a replacement')
      expect(t.slice(fix.start, fix.end)).toBe(fix.match)
      expect(fix.expectedText).toBe(t)
    }
    expect(found[0]?.location).toMatchObject({ kind: 'shape', start: 34, end: 41 })
  })

  it('treats placeholder text as a major item before sending', () => {
    expect(one(wordingFindings(onOneSlide(body('[Insert chart]'))), 'wording/placeholder')).toMatchObject({ severity: 'major', group: 'send' })
  })

  it('reads titles as titles, and labels as compact text', () => {
    const t = 'The World Health Organization report'
    expect(wordingFindings(onOneSlide(title(t)))[0]?.fix).toMatchObject({ replacement: 'Organisation' })
    expect(wordingFindings(onOneSlide(body(t)))[0]?.fix).toBeUndefined()
    expect(wordingFindings(onOneSlide(text('Launch in Jan 2026')))[0]).toMatchObject({ advisory: true })
    expect(wordingFindings(onOneSlide(body('Launch in Jan 2026')))[0]?.fix).toMatchObject({ replacement: 'January 2026' })
  })

  it('offers straight quotes as suggestions, and leaves quotations, furniture and footers alone', () => {
    expect(wordingFindings(onOneSlide(body('He said "yes"'))).map((f) => [f.rule, f.advisory])).toEqual([['wording/quotes', true], ['wording/quotes', true]])
    const skipped = wordingFindings(onOneSlide(
      text('The organization said "yes"', { tags: { BWB_ELEMENT: 'quote' } }),
      text('The organization', { font: font({ name: 'Candara', italic: true }) }),
      text('DRAFT  organization', { tags: { BWB_STICKER: 'x' } }),
      shape({ type: 'Placeholder', placeholder: 'Footer', text: 'The organization' }),
    ))
    expect(skipped).toEqual([])
  })
})
