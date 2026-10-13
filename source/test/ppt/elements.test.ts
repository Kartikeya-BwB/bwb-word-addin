import { afterEach, describe, expect, it } from 'vitest'
import { CONTENT_AREA, SLIDE, SLIDE_TYPE, SOURCE_AREA, TAGS } from '../../src/brand/slideGrid.ts'
import { ALLOWED_HEX } from '../../src/brand/tokens.ts'
import {
  ELEMENTS, insertBwbTable, insertCallout, insertKeyFigures, insertMatrix, insertProcess, insertQuote, insertSourceLine, insertStepMarkers,
  insertTimeline,
} from '../../src/ppt/elements.ts'
import type { ElementSpec } from '../../src/ppt/specs.ts'
import {
  calloutSpec, keyFiguresSpec, markPosition, matrixSpec, noteSpec, processSpec, quoteSpec, ragSpec, sourceLineSpec, stepMarkersSpec,
  stickerSpec, tableSpec, timelineSpec,
} from '../../src/ppt/specs.ts'
import type { FakePresentation, FakeShape } from '../fakePpt/index.ts'
import { allShapes, boundsOf, createPresentation, installFakePowerPoint, slide, textOf, uninstallFakePowerPoint } from '../fakePpt/index.ts'

afterEach(uninstallFakePowerPoint)

function setUp(apiVersion = '1.10'): { deck: FakePresentation; added: () => FakeShape[] } {
  const deck = createPresentation([slide([], { id: 's1' }), slide([], { id: 's2' })], { apiVersion })
  deck.selection.slideIds = ['s2']
  installFakePowerPoint(deck)
  return { deck, added: () => deck.slides[1]?.shapes ?? [] }
}

const fontsOf = (s: FakeShape | undefined): Array<[string, string | null, number | null, string | null, boolean | null]> =>
  s?.textFrame?.runs.map((r) => [r.text, r.font.name, r.font.size, r.font.color, r.font.bold]) ?? []

describe('the element layouts', () => {
  const specs: ElementSpec[] = [
    calloutSpec(), keyFiguresSpec(), quoteSpec(), processSpec(), matrixSpec(true), matrixSpec(false), timelineSpec(), stepMarkersSpec(),
    sourceLineSpec(), noteSpec('13 October 2026', 2),
  ]

  it.each(specs.map((s) => [s.name, s] as const))('%s stays on the slide, in BwB colours and slide type sizes', (_name, spec) => {
    const sizes = new Set<number>([SLIDE_TYPE.title.size, SLIDE_TYPE.body.size, SLIDE_TYPE.small.size, SLIDE_TYPE.source.size])
    for (const s of spec.shapes) {
      const quarter = s.rotation === 270
      const [w, h] = quarter ? [s.height, s.width] : [s.width, s.height]
      const left = s.left + (s.width - w) / 2
      const top = s.top + (s.height - h) / 2
      expect(left, s.name).toBeGreaterThanOrEqual(0)
      expect(top, s.name).toBeGreaterThanOrEqual(0)
      expect(left + w, s.name).toBeLessThanOrEqual(SLIDE.width)
      expect(top + h, s.name).toBeLessThanOrEqual(SLIDE.height)
      expect(s.name).toMatch(/^BwB /)
      for (const colour of [s.fill, s.line?.colour, ...(s.text?.paragraphs.map((p) => p.font.colour) ?? [])]) {
        if (colour) expect(ALLOWED_HEX.has(colour), `${s.name} ${colour}`).toBe(true)
      }
      for (const p of s.text?.paragraphs ?? []) {
        expect(sizes.has(p.font.size), `${s.name} ${p.font.size}`).toBe(true)
        expect(['Aptos', 'Candara']).toContain(p.font.name)
        expect(p.text).not.toMatch(/[\u2014!]/)
      }
    }
  })

  it('puts lime only on navy', () => {
    const quote = quoteSpec().shapes
    const strip = quote.find((s) => s.name === 'BwB quote strip')
    const bar = quote.find((s) => s.name === 'BwB quote bar')
    expect(strip?.fill).toBe('022338')
    expect(bar?.fill).toBe('BED62F')
    expect(bar && strip && bar.left >= strip.left && bar.left + bar.width <= strip.left + strip.width).toBe(true)
    const figures = keyFiguresSpec().shapes
    expect(figures.map((f) => [f.fill, f.text?.paragraphs[0]?.font.colour])).toEqual([['022338', 'BED62F'], ['022338', 'FFFFFF'], ['022338', 'FFFFFF']])
  })

  it('lays the key figures on the 12-column grid and the process across the content area', () => {
    const figures = keyFiguresSpec().shapes
    expect(figures[0]?.left).toBe(CONTENT_AREA.left)
    expect((figures[2]?.left ?? 0) + (figures[2]?.width ?? 0)).toBeCloseTo(CONTENT_AREA.left + CONTENT_AREA.width, 6)
    const steps = processSpec().shapes
    expect(steps.map((s) => s.geometry)).toEqual(['HomePlate', 'Chevron', 'Chevron'])
    expect(steps.map((s) => s.fill)).toEqual(['022338', '022338', '328840'])
    expect((steps[2]?.left ?? 0) + (steps[2]?.width ?? 0)).toBeCloseTo(CONTENT_AREA.left + CONTENT_AREA.width, 6)
  })

  it('sizes a sticker to its words and keeps it in the band above the title', () => {
    const draft = stickerSpec('Draft')
    const long = stickerSpec('Not for distribution')
    expect(long.width).toBeGreaterThan(draft.width)
    expect(draft.left + draft.width).toBe(SLIDE.width - 40)
    expect(draft.top + draft.height).toBeLessThanOrEqual(28)
    expect(draft.text?.paragraphs[0]?.text).toBe('DRAFT')
    expect(stickerSpec('Draft', 720).left + draft.width).toBe(680)
  })

  it('places the next mark beside the last one, wrapping at the margin, or in the middle of the content area', () => {
    expect(markPosition(null, 20)).toEqual({ left: CONTENT_AREA.left + 430, top: CONTENT_AREA.top + 170 })
    expect(markPosition({ left: 100, top: 100, width: 20, height: 20 }, 20)).toEqual({ left: 128, top: 100 })
    expect(markPosition({ left: 900, top: 100, width: 20, height: 20 }, 20)).toEqual({ left: 900, top: 128 })
    expect(ragSpec('amber', { left: 1, top: 2 })).toMatchObject({ geometry: 'Ellipse', fill: 'F2A900', left: 1, top: 2 })
  })

  it('makes a four-by-four table with a header row', () => {
    const t = tableSpec()
    expect(t.values).toHaveLength(4)
    expect(t.values[0]).toEqual(['Heading', 'Heading', 'Heading', 'Heading'])
    expect(t.header).toMatchObject({ fill: '022338', font: { colour: 'FFFFFF', bold: true } })
    expect(t.rule).toEqual({ colour: 'E8E9E6', weight: 0.75 })
  })
})

describe('inserting elements', () => {
  it('adds a callout to the current slide, tagged, named, grouped and selected', async () => {
    const { deck, added } = setUp()
    expect(await insertCallout()).toBe('Callout added at the foot of the content area. Type over the text.')
    const [g] = added()
    expect(g).toMatchObject({ type: 'Group', name: 'BwB callout' })
    expect(g?.tags.get(TAGS.element)).toBe('callout')
    expect(deck.selection.shapeIds).toEqual([g?.id])
    const [bar, box] = g?.members ?? []
    expect(bar?.fill.foregroundColor).toBe('#328840')
    expect(bar?.line.visible).toBe(false)
    expect(box?.fill.foregroundColor).toBe('#F7F8F5')
    expect(box?.tags.get(TAGS.element)).toBe('callout')
    expect(fontsOf(box)).toEqual([['Callout title\n', 'Aptos', 14, '#022338', true], ['One or two sentences with the message.', 'Aptos', 14, '#022338', false]])
    expect(box?.textFrame).toMatchObject({ autoSizeSetting: 'AutoSizeNone', verticalAlignment: 'Middle', leftMargin: 16, wordWrap: true })
    expect(boundsOf(g ?? box ?? bar ?? ({} as FakeShape))).toMatchObject({ left: CONTENT_AREA.left, width: CONTENT_AREA.width })
  })

  it('adds key figures, a quote, a process and a timeline', async () => {
    const { added } = setUp()
    await insertKeyFigures()
    await insertQuote()
    await insertProcess()
    await insertTimeline()
    expect(added().map((s) => [s.name, s.members?.length])).toEqual([['BwB key figures', 3], ['BwB quote', 3], ['BwB process', 3], ['BwB timeline', 9]])
    const figure = added()[0]?.members?.[0]
    expect(fontsOf(figure)[0]).toEqual(['Figure\n', 'Aptos', 24, '#BED62F', true])
    const quoteText = added()[1]?.members?.[2]
    expect(quoteText?.textFrame?.runs[0]?.font).toMatchObject({ name: 'Candara', italic: true, color: '#022338' })
    const steps = added()[2]?.members ?? []
    expect(steps.map((s) => s.geometricShapeType)).toEqual(['HomePlate', 'Chevron', 'Chevron'])
    const timeline = added()[3]?.members ?? []
    expect(timeline[0]).toMatchObject({ type: 'Line', line: { color: '#022338', weight: 2, visible: true } })
  })

  it('turns the matrix axis label when PowerPoint can rotate shapes, and puts it above the matrix when not', async () => {
    const { added } = setUp('1.10')
    await insertMatrix()
    const turned = added()[0]?.members?.find((m) => m.name === 'BwB matrix vertical axis')
    expect(turned?.rotation).toBe(270)
    const older = setUp('1.8')
    await insertMatrix()
    const upright = older.added()[0]?.members?.find((m) => m.name === 'BwB matrix vertical axis')
    expect(upright).toMatchObject({ rotation: 0, top: CONTENT_AREA.top })
  })

  it('leaves the numbered markers as separate shapes, all selected', async () => {
    const { deck, added } = setUp()
    expect(await insertStepMarkers()).toBe('Numbered markers added. Drag each one to its step.')
    expect(added().map((s) => [s.name, textOf(s), s.geometricShapeType])).toEqual([
      ['BwB step marker 1', '1', 'Ellipse'], ['BwB step marker 2', '2', 'Ellipse'], ['BwB step marker 3', '3', 'Ellipse'], ['BwB step marker 4', '4', 'Ellipse'],
    ])
    expect(deck.selection.shapeIds).toEqual(added().map((s) => s.id))
  })

  it('adds one source line per slide, and selects the one there on a second click', async () => {
    const { deck, added } = setUp()
    await insertSourceLine()
    const [line] = added()
    expect(line).toMatchObject({ type: 'TextBox', ...SOURCE_AREA })
    expect(fontsOf(line)).toEqual([['Source: organisation, title (year)', 'Aptos', 8, '#4A4A4C', false]])
    deck.selection.shapeIds = []
    expect(await insertSourceLine()).toBe('This slide has a source line already, so it is selected.')
    expect(added()).toHaveLength(1)
    expect(deck.selection.shapeIds).toEqual([line?.id])
  })

  it('adds the BwB table with a navy header, light grey rules and no PowerPoint style', async () => {
    const { added } = setUp('1.10')
    expect(await insertBwbTable()).toBe('Table added. Type in the cells.')
    const [t] = added()
    expect(t).toMatchObject({ type: 'Table', name: 'BwB table', left: CONTENT_AREA.left, top: CONTENT_AREA.top })
    expect(t?.tags.get(TAGS.element)).toBe('table')
    const options = t?.table?.options as { style?: string; specificCellProperties?: Array<Array<Record<string, unknown>>> } | undefined
    expect(options?.style).toBe('NoStyleNoGrid')
    expect(options?.specificCellProperties?.[0]?.[0]).toMatchObject({ fill: { color: '#022338' }, font: { color: '#FFFFFF', bold: true, name: 'Aptos' } })
    expect(options?.specificCellProperties?.[1]?.[2]).toMatchObject({ fill: { color: '#FFFFFF' }, borders: { top: { color: '#E8E9E6', weight: 0.75 } } })
  })

  it('makes the table without the 1.9 style option on older PowerPoint, and says plainly when it cannot at all', async () => {
    const { added } = setUp('1.8')
    expect(await insertBwbTable()).toMatch(/keeps its own table style/)
    expect((added()[0]?.table?.options as { style?: string }).style).toBeUndefined()
    setUp('1.7')
    await expect(insertBwbTable()).rejects.toThrow('The BwB table needs a newer PowerPoint. Update Microsoft 365, or use PowerPoint on the web.')
  })

  it('leaves the parts ungrouped where PowerPoint cannot group, and says so', async () => {
    const { added } = setUp('1.7')
    expect(await insertCallout()).toMatch(/cannot group shapes from an add-in/)
    expect(added().map((s) => s.name)).toEqual(['BwB callout bar', 'BwB callout text'])
  })

  it('offers every element in the pane, each inserting something', async () => {
    for (const e of ELEMENTS) {
      const { added } = setUp()
      await e.insert()
      expect(allShapes(added()).length, e.id).toBeGreaterThan(0)
    }
  })
})
