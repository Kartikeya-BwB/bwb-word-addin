import { afterEach, describe, expect, it } from 'vitest'
import { takeDeckSnapshot } from '../../../src/ppt/check/snapshot.ts'
import type { ShapeInfo } from '../../../src/ppt/check/types.ts'
import {
  createPresentation, group, installFakePowerPoint, line, picture, placeholder, run, selectSlides, shape, slide, textBox, textFrame,
  uninstallFakePowerPoint,
} from '../../fakePpt/index.ts'

afterEach(uninstallFakePowerPoint)

const byName = (shapes: ShapeInfo[], name: string): ShapeInfo | undefined => shapes.find((s) => s.name === name)

function richDeck() {
  return createPresentation([
    slide([placeholder('CenterTitle', 'Adaptation finance'), placeholder('Subtitle', 'For the board')], { id: 's1', layoutId: 'layout-title', tags: { bwb_layout: 'title' } }),
    slide([
      placeholder('Title', 'Growth slowed in 2025', { id: 't2' }),
      placeholder('Content', 'Point one\rPoint two', { id: 'c2' }),
      textBox('Source: BwB analysis', { id: 'src', name: 'Source', top: 480, font: { size: 8, color: '#4a4a4c' } }),
      picture({ id: 'pic', altTextDescription: 'Solar panels on a roof' }),
      shape({ id: 'rag', fill: { type: 'Solid', foregroundColor: '#F2A900', transparency: 0 }, tags: { BWB_RAG: 'amber' } }),
      line({ id: 'rule' }),
    ], { id: 's2', tags: { BWB_SECTION: 'Context' } }),
  ], { pageSetup: { slideWidth: 720, slideHeight: 540 } })
}

describe('reading the deck', () => {
  it('reads slides, shapes, text, fonts, colours, geometry, placeholders, tags, alt text and the slide size', async () => {
    const deck = richDeck()
    selectSlides(deck, ['s2'])
    installFakePowerPoint(deck)
    const snap = await takeDeckSnapshot()
    expect(snap.slides.map((s) => [s.id, s.index, s.layout, s.tags])).toEqual([
      ['s1', 0, 'Title Slide', { BWB_LAYOUT: 'title' }],
      ['s2', 1, 'Title and Content', { BWB_SECTION: 'Context' }],
    ])
    expect([snap.width, snap.height, snap.selected, snap.partial]).toEqual([720, 540, ['s2'], false])
    expect(snap.can).toEqual({
      formatting: true, placeholders: true, groups: true, altText: true, pageSize: true, selection: true, hidden: true, background: true, links: true, markers: true,
    })

    const shapes = snap.slides[1]?.shapes ?? []
    expect(shapes.map((s) => [s.id, s.type, s.placeholder])).toEqual([
      ['t2', 'Placeholder', 'Title'], ['c2', 'Placeholder', 'Content'], ['src', 'TextBox', null], ['pic', 'Image', null],
      ['rag', 'GeometricShape', null], ['rule', 'Line', null],
    ])
    const title = shapes[0]
    expect(title).toMatchObject({ path: ['t2'], level: 0, text: 'Growth slowed in 2025', left: 40, top: 28, width: 880, height: 64 })
    expect(title?.font).toEqual({ name: 'Candara', size: 24, colour: '000000', bold: true, italic: false })
    expect(byName(shapes, 'Source')?.font).toMatchObject({ size: 8, colour: '4A4A4C' })
    expect(byName(shapes, 'Source')?.fill).toEqual({ type: 'NoFill', colour: 'FFFFFF' })
    expect(shapes[3]).toMatchObject({ text: null, font: null, fill: null, alt: { title: '', description: 'Solar panels on a roof', decorative: false } })
    expect(shapes[4]).toMatchObject({ tags: { BWB_RAG: 'amber' }, fill: { type: 'Solid', colour: 'F2A900' }, line: { visible: true, colour: '042433' } })
    expect(shapes[5]).toMatchObject({ text: null, line: { visible: true, colour: '156082' } })
  })

  it('reads only text and tags for the slide tools, and only the slides asked for', async () => {
    const deck = richDeck()
    installFakePowerPoint(deck)
    const light = await takeDeckSnapshot({ formatting: false })
    const title = light.slides[1]?.shapes[0]
    expect(title).toMatchObject({ text: 'Growth slowed in 2025', font: null, fill: null, line: null })
    expect(light.slides[1]?.shapes[3]?.alt).toBeNull()
    expect(light.can.formatting).toBe(false)
    const one = await takeDeckSnapshot({ slideIds: ['s2'] })
    expect(one.slides.map((s) => [s.id, s.index])).toEqual([['s2', 1]])
    selectSlides(deck, ['s1'])
    const current = await takeDeckSnapshot({ current: true, formatting: false })
    expect(current.slides.map((s) => s.id)).toEqual(['s1'])
  })

  it('splits mixed text into paragraphs, then words, reading each piece', async () => {
    const mixed = textBox('', {
      id: 'mixed',
      textFrame: textFrame('', { runs: [run('Hello ', { name: 'Aptos' }), run('world', { name: 'Arial' }), run('\rSecond line', { name: 'Aptos', size: 9 })] }),
    })
    const oneLine = textBox('', { id: 'line', textFrame: textFrame('', { runs: [run('Red', { color: '#FF0000' }), run(' and black')] }) })
    installFakePowerPoint(createPresentation([slide([mixed, oneLine])]))
    const shapes = (await takeDeckSnapshot()).slides[0]?.shapes ?? []
    const font = (name: string, size = 18, colour = '000000') => ({ name, size, colour, bold: false, italic: false })
    expect(shapes[0]?.font).toMatchObject({ name: null, size: null, colour: '000000' })
    expect(shapes[0]?.pieces).toEqual([
      { start: 0, length: 5, font: font('Aptos') },
      { start: 6, length: 5, font: font('Arial') },
      { start: 12, length: 11, font: font('Aptos', 9) },
    ])
    expect(shapes[1]?.pieces).toEqual([
      { start: 0, length: 3, font: font('Aptos', 18, 'FF0000') },
      { start: 4, length: 3, font: font('Aptos') },
      { start: 8, length: 5, font: font('Aptos') },
    ])
  })

  it('reads the shapes inside groups, a level at a time, with their paths', async () => {
    const inner = group([textBox('Deep', { id: 'deep', left: 300 }), shape({ id: 'box', left: 400 })], { id: 'inner' })
    const outer = group([textBox('Inside', { id: 'inside', font: { name: 'Arial' } }), inner], { id: 'outer' })
    installFakePowerPoint(createPresentation([slide([outer, textBox('Top', { id: 'top' })])]))
    const shapes = (await takeDeckSnapshot()).slides[0]?.shapes ?? []
    expect(shapes.map((s) => [s.id, s.level, s.path.join('/')])).toEqual([
      ['outer', 0, 'outer'], ['inside', 1, 'outer/inside'], ['inner', 1, 'outer/inner'],
      ['deep', 2, 'outer/inner/deep'], ['box', 2, 'outer/inner/box'], ['top', 0, 'top'],
    ])
    expect(shapes.find((s) => s.id === 'inside')?.font?.name).toBe('Arial')
    expect(shapes.find((s) => s.id === 'outer')).toMatchObject({ text: null, left: 0, width: 500 })
  })

  it('keeps the number of syncs the same however many slides there are', async () => {
    const deck = (n: number) => createPresentation(Array.from({ length: n }, (_, i) => slide([
      placeholder('Title', `Slide ${i}`), placeholder('Content', 'Text'), group([textBox('a'), textBox('b')]),
    ])))
    const small = installFakePowerPoint(deck(2))
    await takeDeckSnapshot()
    const few = small.syncs()
    uninstallFakePowerPoint()
    const large = installFakePowerPoint(deck(30))
    await takeDeckSnapshot()
    expect(large.syncs()).toBe(few)
    expect(few).toBeLessThanOrEqual(10)
  })

  it('reads alt text on graphics such as icons, and whether a shape is decorative', async () => {
    const icon = picture({ id: 'icon', type: 'Graphic', name: 'Graphic 4' })
    const decorative = picture({ id: 'line-art', type: 'Graphic', isDecorative: true })
    installFakePowerPoint(createPresentation([slide([icon, decorative])]))
    const shapes = (await takeDeckSnapshot()).slides[0]?.shapes ?? []
    expect(shapes.map((s) => s.alt)).toEqual([
      { title: '', description: '', decorative: false }, { title: '', description: '', decorative: true },
    ])
  })
})

describe('reading the deck on older PowerPoint', () => {
  it('on API 1.5, finds placeholders by name and leaves groups, alt text and the slide size alone', async () => {
    const deck = richDeck()
    deck.slides[1]?.shapes.push(group([textBox('Inside'), textBox('Also')], { id: 'g' }))
    installFakePowerPoint(deck, { apiVersion: '1.5' })
    const snap = await takeDeckSnapshot()
    expect(snap.can).toEqual({
      formatting: true, placeholders: false, groups: false, altText: false, pageSize: false, selection: true, hidden: false, background: false, links: false, markers: false,
    })
    expect([snap.width, snap.height]).toEqual([960, 540])
    const shapes = snap.slides[1]?.shapes ?? []
    expect(shapes.map((s) => s.placeholder).slice(0, 2)).toEqual(['Title', 'Body'])
    expect(shapes.find((s) => s.id === 'pic')?.alt).toBeNull()
    expect(shapes.filter((s) => s.level > 0)).toEqual([])
    expect(snap.slides[0]?.shapes[0]?.placeholder).toBe('Title')
  })

  it('on API 1.4, cannot say which slide is current', async () => {
    installFakePowerPoint(richDeck(), { apiVersion: '1.4' })
    const snap = await takeDeckSnapshot()
    expect(snap.selected).toEqual([])
    expect(snap.can.selection).toBe(false)
    expect(snap.slides).toHaveLength(2)
  })

  it('says plainly when PowerPoint is too old to read the slides, or missing', async () => {
    installFakePowerPoint(richDeck(), { apiVersion: '1.3' })
    await expect(takeDeckSnapshot()).rejects.toThrow('Reading the slides needs a newer PowerPoint')
    uninstallFakePowerPoint()
    await expect(takeDeckSnapshot()).rejects.toThrow('Open this in PowerPoint to use it.')
  })

  it('reads the rest of the deck when one shape cannot be read', async () => {
    // A content placeholder holding a picture has no text frame; before API 1.8 nothing says so in advance.
    const holdsPicture = placeholder('Content', '', { id: 'holds', textFrame: null })
    const deck = createPresentation([slide([placeholder('Title', 'A title', { id: 't' }), holdsPicture, textBox('Body text', { id: 'b' })])])
    installFakePowerPoint(deck, { apiVersion: '1.5' })
    const shapes = (await takeDeckSnapshot()).slides[0]?.shapes ?? []
    expect(shapes.map((s) => [s.id, s.text])).toEqual([['t', 'A title'], ['holds', null], ['b', 'Body text']])
    expect(shapes.find((s) => s.id === 'b')?.font?.name).toBe('Aptos')
  })
})
