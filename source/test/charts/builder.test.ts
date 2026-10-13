import { afterEach, describe, expect, it, vi } from 'vitest'
import { DEFAULT_SIZES, renderChart } from '../../src/charts/spec.ts'
import {
  initialState, placement, sizeOptions, toSpec, uses, withExample, withGrid, withPaste, withType,
} from '../../src/ui/charts/builderState.ts'
import { MAX_RECENT, RECENT_KEY, loadRecent, recentLabel, saveRecent } from '../../src/ui/charts/recent.ts'
import { insertChart, pixelSize, sizedSvg } from '../../src/ui/charts/insertChart.ts'
import { createModel, para } from '../fakeWord/model.ts'
import { installFakeWord, uninstallFakeWord } from '../fakeWord/install.ts'

vi.mock('../../src/word/iconRenderer.ts', () => ({
  svgToPng: vi.fn(async () => new Uint8Array([137, 80, 78, 71, 13, 10, 26, 10])),
}))

const TODAY = '2026-10-13'

describe('builder state', () => {
  it('starts from the column example at the Word default size', () => {
    const spec = toSpec(initialState('word'), 'word', TODAY)
    expect(spec.type).toBe('column')
    expect(spec.size).toEqual(DEFAULT_SIZES.word)
    expect(spec.fontPt).toBe(9)
    expect(spec.title).toBe('Funding by year')
    expect(spec.sourceInChart).toBe(false)
    expect(spec.format).toEqual({ currency: 'USD', scale: 'm', percent: false })
  })

  it('starts in PowerPoint at the content area, with the source line in the chart', () => {
    const state = initialState('powerpoint')
    expect(placement(state, 'powerpoint')).toEqual({ widthPt: 880, heightPt: 360, fontPt: 12, left: 40, top: 108 })
    expect(toSpec(state, 'powerpoint', TODAY).sourceInChart).toBe(true)
  })

  it('offers full, half and custom sizes in each host', () => {
    expect(sizeOptions('word').map((o) => o.value)).toEqual(['full', 'half', 'custom'])
    const half = placement({ ...initialState('powerpoint'), size: 'half' }, 'powerpoint')
    expect(half).toMatchObject({ widthPt: 432, heightPt: 360, left: 40, top: 108 })
    const word = placement({ ...initialState('word'), size: 'half' }, 'word')
    expect(word.widthPt).toBeCloseTo(232.44, 1)
    const custom = placement({ ...initialState('word'), size: 'custom', customWidth: '40', customHeight: 'x' }, 'word')
    expect(custom.widthPt).toBeCloseTo(470.55, 1)
    expect(custom.heightPt).toBeCloseTo(255.12, 1)
    const slide = placement({ ...initialState('powerpoint'), size: 'custom', customWidth: '400', customHeight: '200' }, 'powerpoint')
    expect(slide).toEqual({ widthPt: 400, heightPt: 200, fontPt: 11, left: 280, top: 188 })
  })

  it('swaps in the next example while the data is still an example, and keeps the author\u{2019}s own data', () => {
    const gantt = withType(initialState('word'), 'gantt')
    expect(gantt.grid[0]).toEqual(['Task', 'Start', 'End'])
    expect(gantt.title).toBe('Workplan')
    const own = withGrid(initialState('word'), [['K', 'V'], ['A', '1']])
    expect(own.title).toBe('')
    expect(withType(own, 'bar').grid).toEqual([['K', 'V'], ['A', '1']])
    expect(withExample(own).grid).toEqual(initialState('word').grid)
  })

  it('takes units from pasted cells, and keeps the author\u{2019}s units for plain numbers', () => {
    const pasted = withPaste(initialState('word'), 'Country\tSize\nKenya\tUSD 1.2bn\nGhana\tUSD 900m\n')
    expect(pasted).toMatchObject({ units: 'bn', currency: 'USD', example: false, title: '', source: '' })
    expect(pasted.grid).toEqual([['Country', 'Size'], ['Kenya', 'USD 1.2bn'], ['Ghana', 'USD 900m']])
    const typed = withPaste({ ...pasted, units: 'k' }, 'Country\tSize\nKenya\t5')
    expect(typed.units).toBe('k')
    expect(withPaste(initialState('word'), 'A\tB\n1\t2').units).toBe('')
  })

  it('passes only the options a type uses', () => {
    const state = { ...initialState('word'), axisMin: '10', axisMax: '', highlight: '2024', today: true }
    expect(toSpec(state, 'word', TODAY).axis).toEqual({ min: 10 })
    expect(toSpec(state, 'word', TODAY).highlight).toBe('2024')
    const donut = toSpec({ ...state, type: 'mekko' }, 'word', TODAY)
    expect(donut.axis).toBeUndefined()
    expect(donut.highlight).toBeUndefined()
    expect(toSpec(state, 'word', TODAY).today).toBe(TODAY)
    expect(uses('gantt', 'format')).toBe(false)
    expect(uses('line', 'markers')).toBe(true)
  })

  it('maps the number format fields', () => {
    const spec = toSpec({ ...initialState('word'), units: '%', currency: 'eur', decimals: '1' }, 'word', TODAY)
    expect(spec.format).toEqual({ currency: 'EUR', scale: '', percent: true, decimals: 1 })
  })
})

function memoryStore(): Storage & { data: Map<string, string> } {
  const data = new Map<string, string>()
  return {
    data,
    get length() { return data.size },
    clear: () => data.clear(),
    getItem: (k: string) => data.get(k) ?? null,
    key: (i: number) => [...data.keys()][i] ?? null,
    removeItem: (k: string) => void data.delete(k),
    setItem: (k: string, v: string) => void data.set(k, v),
  }
}

describe('recent charts', () => {
  it('keeps the last ten, newest first, without duplicates', () => {
    const store = memoryStore()
    for (let i = 0; i < 12; i += 1) saveRecent({ ...initialState('word'), title: `Chart ${i}` }, 'word', store)
    saveRecent({ ...initialState('word'), title: 'Chart 11' }, 'word', store)
    const list = loadRecent('word', store)
    expect(list).toHaveLength(MAX_RECENT)
    expect(list.map(recentLabel).slice(0, 3)).toEqual(['Chart 11', 'Chart 10', 'Chart 9'])
  })

  it('reopens a chart with fields added since it was saved', () => {
    const store = memoryStore()
    const { colours: _colours, ...older } = initialState('word')
    store.setItem(RECENT_KEY, JSON.stringify([{ id: 'a', savedAt: '2026-10-13T09:00:00Z', state: { ...older, title: 'Old' } }]))
    const [chart] = loadRecent('word', store)
    expect(chart?.state.title).toBe('Old')
    expect(chart?.state.colours).toEqual([])
    expect(renderChart(toSpec(chart?.state ?? initialState('word'), 'word', TODAY)).svg).toContain('<svg')
  })

  it('ignores damaged or blocked storage', () => {
    const store = memoryStore()
    store.setItem(RECENT_KEY, '{not json')
    expect(loadRecent('word', store)).toEqual([])
    store.setItem(RECENT_KEY, JSON.stringify([{ id: 'x', savedAt: 'y', state: { type: 'pie', grid: [['a']] } }, 7]))
    expect(loadRecent('word', store)).toEqual([])
    const blocked = { getItem: () => { throw new Error('denied') }, setItem: () => { throw new Error('full') } }
    expect(loadRecent('word', blocked)).toEqual([])
    expect(() => saveRecent(initialState('word'), 'word', blocked)).not.toThrow()
    expect(loadRecent('word', null)).toEqual([])
  })
})

describe('insertChart', () => {
  afterEach(() => {
    uninstallFakeWord()
    delete (globalThis as Record<string, unknown>).Office
  })

  it('sizes the picture at four times 96 dpi', () => {
    expect(pixelSize({ widthPt: 470.55, heightPt: 255.12 })).toEqual({ width: 2510, height: 1361 })
    expect(sizedSvg('<svg width="1" height="2" viewBox="0 0 3 4">', 10, 20)).toBe('<svg viewBox="0 0 3 4" width="10" height="20">')
  })

  it('puts a titled chart in Word as one figure block with its numbered caption and alt text', async () => {
    const empty = para('')
    const model = createModel([para('Before'), empty], { selection: [empty] })
    installFakeWord(model)
    const chart = renderChart(toSpec(initialState('word'), 'word', TODAY))
    const message = await insertChart(chart, { host: 'word', title: 'Funding by year', source: 'BwB analysis (2026)', sourceInChart: false })
    expect(message).toMatch(/caption/)
    expect(model.inserts).toHaveLength(1)
    const ooxml = model.inserts[0]?.ooxml ?? ''
    expect(ooxml).toContain('Column chart: Funding by year')
    expect(ooxml).toMatch(/SEQ Figure/)
    expect(ooxml).toContain('Funding by year')
    expect(ooxml).toContain('Source: BwB analysis (2026)')
    expect(ooxml).toMatch(/<pic:pic|<wp:inline/)
  })

  it('leaves the source out of the caption when the chart draws it', async () => {
    const empty = para('')
    const model = createModel([empty], { selection: [empty] })
    installFakeWord(model)
    const chart = renderChart(toSpec({ ...initialState('word'), sourceInChart: true }, 'word', TODAY))
    await insertChart(chart, { host: 'word', title: 'Funding', source: 'BwB analysis (2026)', sourceInChart: true })
    expect(model.inserts[0]?.ooxml).toMatch(/SEQ Figure/)
    expect(model.inserts[0]?.ooxml).not.toContain('Source:')
  })

  it('puts the SVG on the slide at the content area in PowerPoint', async () => {
    const calls: Array<{ data: string; options: Record<string, unknown> }> = []
    ;(globalThis as Record<string, unknown>).Office = {
      context: {
        host: 'PowerPoint',
        document: {
          setSelectedDataAsync: (data: string, options: Record<string, unknown>, done: (r: { status: string }) => void) => {
            calls.push({ data, options })
            done({ status: 'succeeded' })
          },
        },
      },
      HostType: { PowerPoint: 'PowerPoint', Word: 'Word' },
      CoercionType: { XmlSvg: 'xmlSvg', Image: 'image' },
      AsyncResultStatus: { Succeeded: 'succeeded', Failed: 'failed' },
    }
    const state = initialState('powerpoint')
    const chart = renderChart(toSpec(state, 'powerpoint', TODAY))
    await insertChart(chart, { host: 'powerpoint', title: '', source: '', sourceInChart: true, left: 40, top: 108 })
    expect(calls).toHaveLength(1)
    expect(calls[0]?.data.startsWith('<svg')).toBe(true)
    expect(calls[0]?.options).toMatchObject({ coercionType: 'xmlSvg', imageWidth: 880, imageHeight: 360, imageLeft: 40, imageTop: 108 })
  })
})
