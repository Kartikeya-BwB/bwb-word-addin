import { describe, expect, it } from 'vitest'
import { CHART_HEX } from '../../src/charts/palette.ts'
import { SAMPLES } from '../../src/charts/samples.ts'
import type { ChartSpec, ChartType } from '../../src/charts/spec.ts'
import { CHART_TYPES, DEFAULT_SIZES, defaultOptions, renderChart } from '../../src/charts/spec.ts'
import { colours, overlaps, parseSvg, textBoxes } from './svgCheck.ts'

const HALF_WORD = { widthPt: 232, heightPt: 255 }

function sampleSpec(type: ChartType, size = DEFAULT_SIZES.word, extra: Partial<ChartSpec> = {}): ChartSpec {
  const sample = SAMPLES[type]
  return {
    ...defaultOptions(type),
    type,
    data: sample.data,
    title: sample.title,
    source: sample.source,
    size,
    ...(sample.format ? { format: sample.format } : {}),
    ...(sample.highlight ? { highlight: sample.highlight } : {}),
    ...(type === 'gantt' ? { today: '2027-01-18' } : {}),
    ...extra,
  }
}

const cases: Array<[string, ChartSpec]> = []
for (const { type } of CHART_TYPES) {
  cases.push([`${type} in Word`, sampleSpec(type)])
  cases.push([`${type} in PowerPoint with the source drawn`, sampleSpec(type, DEFAULT_SIZES.powerpoint, { sourceInChart: true })])
  cases.push([`${type} at half width in Word`, sampleSpec(type, HALF_WORD)])
  cases.push([`${type} with value axis and no labels`, sampleSpec(type, DEFAULT_SIZES.word, { labels: false, valueAxis: true, legend: true })])
}

const extraCases: Array<[string, ChartSpec]> = [
  ['columns with negative values', {
    type: 'column', size: DEFAULT_SIZES.word,
    data: [['Year', 'Net flow'], ['2021', '-12.5'], ['2022', '8'], ['2023', '-3.25'], ['2024', '15'], ['2025', '21.75']],
  }],
  ['clustered columns, three series and the value axis', {
    type: 'column', size: DEFAULT_SIZES.word, valueAxis: true,
    data: [['Year', 'Grants', 'Loans', 'Equity'], ['2022', 'USD 12m', 'USD 40m', 'USD 5m'], ['2023', 'USD 15m', 'USD 46m', 'USD 7m'], ['2024', 'USD 19m', 'USD 51m', 'USD 9m']],
  }],
  ['twelve months of columns', {
    type: 'column', size: DEFAULT_SIZES.word,
    data: [['Month', 'Disbursed'], ...['January', 'February', 'March', 'April', 'May', 'June', 'July', 'August', 'September', 'October', 'November', 'December'].map((m, i) => [m, `${(i + 1) * 1234.5}`])],
  }],
  ['bars with long names and negatives', {
    type: 'bar', size: DEFAULT_SIZES.word,
    data: [['Fund', 'Return'], ['Emerging markets sovereign debt fund', '4.2%'], ['Global green infrastructure partnership', '-1.8%'], ['Frontier private equity', '7.5%']],
  }],
  ['a line with gaps and one series', {
    type: 'line', size: DEFAULT_SIZES.word,
    data: [['Year', 'Yield'], ['2019', '2.1%'], ['2020', '1.4%'], ['2021', ''], ['2022', '3.9%'], ['2023', '4.6%'], ['2024', '4.1%']],
  }],
  ['a donut with eight segments, folded into six', {
    type: 'donut', size: DEFAULT_SIZES.word,
    data: [['Sector', 'Share'], ...['A', 'B', 'C', 'D', 'E', 'F', 'G', 'H'].map((s, i) => [s, `${20 - i * 2}`])],
  }],
  ['a waterfall crossing zero', {
    type: 'waterfall', size: DEFAULT_SIZES.powerpoint,
    data: [['Step', 'Value', 'Kind'], ['Opening', '10', 'total'], ['Loss', '-25', ''], ['Grant', '8', ''], ['Closing', '', 'total']],
  }],
  ['a Gantt over three years, in quarters', {
    type: 'gantt', size: DEFAULT_SIZES.powerpoint,
    data: [['Task', 'Start', 'End'], ['Phase 1', 'Q1 2026', 'Q4 2026'], ['Phase 2', 'Q3 2026', 'Q2 2028'], ['Review', 'H1 2027', ''], ['Phase 3', '2028-01-01', '2028-12-31']],
  }],
]

describe.each([...cases, ...extraCases])('%s', (_, spec) => {
  const chart = renderChart(spec)
  const parsed = parseSvg(chart.svg)

  it('is well-formed SVG sized in points', () => {
    expect(parsed.errors).toEqual([])
    expect(parsed.root.tagName).toBe('svg')
    expect(parsed.root.getAttribute('xmlns')).toBe('http://www.w3.org/2000/svg')
    expect(parsed.width).toBeCloseTo(spec.size.widthPt, 1)
    expect(parsed.height).toBeCloseTo(spec.size.heightPt, 1)
    expect(chart.widthPt).toBe(spec.size.widthPt)
  })

  it('uses only BwB chart colours, and never lime', () => {
    const used = colours(parsed.root)
    for (const colour of used) expect(CHART_HEX.has(colour), `unexpected colour ${colour}`).toBe(true)
    expect(used.has('BED62F')).toBe(false)
  })

  it('sets no text below 8 pt', () => {
    const boxes = textBoxes(parsed.root)
    expect(boxes.length).toBeGreaterThan(0)
    for (const box of boxes) expect(box.size, box.text).toBeGreaterThanOrEqual(8)
  })

  it('keeps labels apart and inside the chart', () => {
    const boxes = textBoxes(parsed.root)
    expect(overlaps(boxes)).toEqual([])
    for (const box of boxes) {
      expect(box.x0, box.text).toBeGreaterThanOrEqual(-0.5)
      expect(box.x1, box.text).toBeLessThanOrEqual(parsed.width + 0.5)
      expect(box.y0, box.text).toBeGreaterThanOrEqual(-0.5)
      expect(box.y1, box.text).toBeLessThanOrEqual(parsed.height + 0.5)
    }
  })

  it('draws no title, borders or shadows', () => {
    expect(chart.svg).not.toMatch(/<title|filter|<rect[^>]*stroke=|opacity/)
    if (spec.title) expect(textBoxes(parsed.root).some((b) => b.text === spec.title)).toBe(false)
  })
})

describe('escaping', () => {
  it('escapes every piece of text, so names with markup characters stay text', () => {
    const name = 'R&D <pilot> "A" \u{2019}s'
    const chart = renderChart({
      type: 'column', size: DEFAULT_SIZES.word, title: 'Spend & <growth>',
      data: [['Item', `Cost & "fees"`, 'Other'], [name, '5', '1'], ['Plain', '7', '2']],
    })
    const parsed = parseSvg(chart.svg)
    expect(parsed.errors).toEqual([])
    const texts = textBoxes(parsed.root).map((b) => b.text)
    expect(texts).toContain('Cost & "fees"')
    expect(texts.some((t) => t.startsWith('R&D'))).toBe(true)
    expect(chart.svg).not.toContain('<pilot>')
    expect(parsed.root.getAttribute('aria-label')).toBe(chart.altText)
    expect(chart.altText).toContain('Spend & <growth>')
  })

  it('drops characters XML cannot hold', () => {
    const chart = renderChart({ type: 'bar', size: DEFAULT_SIZES.word, data: [['K', 'V'], ['Bad\u{0001}name', '3']] })
    expect(parseSvg(chart.svg).errors).toEqual([])
  })
})

describe('alt text', () => {
  it('names the type, the title and the span', () => {
    const chart = renderChart(sampleSpec('column'))
    expect(chart.altText.startsWith('Column chart: Funding by year, 2022 to 2025.')).toBe(true)
    expect(chart.altText).toContain('2025: USD 240m')
  })

  it('describes every type', () => {
    const alt = (type: ChartType): string => renderChart(sampleSpec(type)).altText
    expect(alt('stackedColumn')).toMatch(/^Stacked column chart: .*, 2021 to 2025\. Series: Grants, Loans, Equity and Guarantees\./)
    expect(alt('column100')).toMatch(/^100% stacked column chart: /)
    expect(alt('bar')).toContain('Germany: USD 72.3bn')
    expect(alt('line')).toMatch(/2019 to 2025/)
    expect(alt('waterfall')).toContain('From Budget 2025 (100) to Budget 2026 (109), in 4 changes.')
    expect(alt('mekko')).toMatch(/^Mekko chart: /)
    expect(alt('gantt')).toMatch(/^Gantt chart: Workplan, November 2026 to April 2027\. 4 tasks and 4 milestones/)
    expect(alt('donut')).toContain('Energy USD 38m (38%)')
    expect(alt('harvey')).toContain('4 options scored from 0 to 4 against 5 criteria')
  })

  it('works without a title', () => {
    expect(renderChart({ type: 'line', size: DEFAULT_SIZES.word, data: SAMPLES.line.data }).altText).toMatch(/^Line chart, 2019 to 2025\./)
  })
})

describe('options', () => {
  it('highlights one category in green and the rest in navy', () => {
    const chart = renderChart(sampleSpec('bar', DEFAULT_SIZES.word, { highlight: 'France' }))
    const fills = [...chart.svg.matchAll(/<rect[^>]*fill="#([0-9A-F]{6})"/g)].map((m) => m[1])
    expect(fills.filter((f) => f === '328840')).toHaveLength(1)
    expect(fills.filter((f) => f === '022338')).toHaveLength(5)
  })

  it('sorts by value when asked', () => {
    const chart = renderChart({ type: 'column', size: DEFAULT_SIZES.word, sort: 'desc', data: [['K', 'V'], ['Low', '1'], ['High', '9'], ['Mid', '5']] })
    const labels = textBoxes(parseSvg(chart.svg).root).filter((b) => ['Low', 'High', 'Mid'].includes(b.text)).sort((a, b) => a.x0 - b.x0)
    expect(labels.map((b) => b.text)).toEqual(['High', 'Mid', 'Low'])
  })

  it('turns labels and the legend off', () => {
    const chart = renderChart(sampleSpec('stackedColumn', DEFAULT_SIZES.word, { labels: false, legend: false }))
    const texts = textBoxes(parseSvg(chart.svg).root).map((b) => b.text)
    expect(texts).not.toContain('Grants')
    expect(texts).not.toContain('120.2')
  })

  it('keeps the author\u{2019}s axis range', () => {
    const chart = renderChart({ type: 'line', size: DEFAULT_SIZES.word, valueAxis: true, axis: { min: 0, max: 1000 }, data: SAMPLES.line.data })
    const texts = textBoxes(parseSvg(chart.svg).root).map((b) => b.text)
    expect(texts).toContain('1,000')
    expect(texts).toContain('0')
  })

  it('writes the units once when the value axis shows numbers', () => {
    const chart = renderChart({ type: 'column', size: DEFAULT_SIZES.word, valueAxis: true, data: SAMPLES.column.data })
    const texts = textBoxes(parseSvg(chart.svg).root).map((b) => b.text)
    expect(texts).toContain('USD m')
    expect(texts).toContain('240')
    expect(texts.some((t) => t.startsWith('USD 2'))).toBe(false)
  })

  it('draws the source line only when asked, in 8 pt', () => {
    const without = renderChart(sampleSpec('column'))
    expect(without.svg).not.toContain('Source:')
    const withSource = renderChart(sampleSpec('column', DEFAULT_SIZES.word, { sourceInChart: true }))
    const source = textBoxes(parseSvg(withSource.svg).root).find((b) => b.text.startsWith('Source:'))
    expect(source?.text).toBe('Source: BwB analysis (2026)')
    expect(source?.size).toBe(8)
  })

  it('reports what it could not draw', () => {
    expect(renderChart({ type: 'column', size: DEFAULT_SIZES.word, data: [['K', 'V'], ['A', 'x']] }).warnings.length).toBeGreaterThan(0)
    expect(renderChart({ type: 'harvey', size: DEFAULT_SIZES.word, data: [['O', 'C'], ['A', '7']] }).warnings).toHaveLength(1)
  })
})

describe('axis range', () => {
  it('clips bars to an axis range the author set, so nothing runs into the labels', () => {
    for (const type of ['column', 'bar'] as const) {
      const chart = renderChart({
        type, size: DEFAULT_SIZES.word, axis: { min: 50, max: 150 }, valueAxis: true,
        data: [['K', 'V'], ['Low', '20'], ['Mid', '100'], ['High', '400']],
      })
      const parsed = parseSvg(chart.svg)
      expect(overlaps(textBoxes(parsed.root))).toEqual([])
      const rects = [...chart.svg.matchAll(/<rect x="([\d.]+)" y="([\d.]+)" width="([\d.]+)" height="([\d.]+)"/g)].map((m) => m.slice(1).map(Number))
      const boxes = textBoxes(parsed.root)
      for (const [x = 0, y = 0, w = 0, h = 0] of rects) {
        expect(x).toBeGreaterThanOrEqual(0)
        expect(y).toBeGreaterThanOrEqual(0)
        for (const b of boxes) {
          const apart = b.x1 <= x + 0.25 || b.x0 >= x + w - 0.25 || b.y1 <= y + 0.25 || b.y0 >= y + h - 0.25
          expect(apart, `${b.text} under a bar`).toBe(true)
        }
      }
    }
  })
})

describe('negative stacks', () => {
  it('puts the total of an all-negative stack beyond its far end, clear of everything else', () => {
    const data = [['Year', 'A', 'B'], ['2024', '-5', '-3'], ['2025', '4', '6'], ['2026', '-2', '7']]
    for (const type of ['stackedColumn', 'stackedBar'] as const) {
      const chart = renderChart({ type, size: DEFAULT_SIZES.word, data })
      const parsed = parseSvg(chart.svg)
      const boxes = textBoxes(parsed.root)
      expect(overlaps(boxes)).toEqual([])
      expect(boxes.some((b) => b.text === '\u{2212}8')).toBe(true)
      for (const box of boxes) {
        expect(box.y0).toBeGreaterThanOrEqual(-0.5)
        expect(box.y1).toBeLessThanOrEqual(parsed.height + 0.5)
      }
    }
  })
})
