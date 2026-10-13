// Visual check: draws every chart type with its sample data at the Word and PowerPoint default sizes, and writes the
// SVG and a PNG of each. Run it with Node 22.18 or later, which runs TypeScript directly:
//
//   node test/charts/render-samples.ts [output folder]
//
// The PNGs use the fonts installed on the machine (Aptos and Candara when present), like Office does.

import { mkdirSync, writeFileSync } from 'node:fs'
import { tmpdir } from 'node:os'
import { join } from 'node:path'
import { Resvg } from '@resvg/resvg-js'
import { SAMPLES } from '../../src/charts/samples.ts'
import type { ChartSpec } from '../../src/charts/spec.ts'
import { CHART_TYPES, DEFAULT_SIZES, defaultOptions, renderChart } from '../../src/charts/spec.ts'

const out = process.argv[2] ?? join(tmpdir(), 'bwb-charts')
mkdirSync(out, { recursive: true })

for (const { type } of CHART_TYPES) {
  const sample = SAMPLES[type]
  for (const host of ['word', 'powerpoint'] as const) {
    const spec: ChartSpec = {
      ...defaultOptions(type),
      type,
      data: sample.data,
      title: sample.title,
      source: sample.source,
      size: DEFAULT_SIZES[host],
      sourceInChart: host === 'powerpoint' && Boolean(sample.source),
      ...(type === 'gantt' ? { today: '2027-01-18' } : {}),
      ...(sample.format ? { format: sample.format } : {}),
      ...(sample.highlight ? { highlight: sample.highlight } : {}),
    }
    const chart = renderChart(spec)
    const name = `${type}-${host}`
    writeFileSync(join(out, `${name}.svg`), chart.svg)
    const scale = host === 'word' ? 2 : 1.5
    const png = new Resvg(chart.svg, {
      fitTo: { mode: 'width', value: Math.round(((chart.widthPt * 4) / 3) * scale) },
      background: 'white',
      font: { loadSystemFonts: true, defaultFontFamily: 'Aptos' },
    }).render().asPng()
    writeFileSync(join(out, `${name}.png`), png)
    const notes = chart.warnings.length ? `  warnings: ${chart.warnings.join(' | ')}` : ''
    console.log(`${name}.png  ${chart.altText}${notes}`)
  }
}
console.log(`Written to ${out}`)
