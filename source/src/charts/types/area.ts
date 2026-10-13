// Stacked area and 100% area charts: categories (usually periods) along the bottom, one band per series stacked from
// the bottom in series order, with thin white lines between bands. With labels on, each band's last value (or share)
// is written beyond the right-hand end, spread so the labels never touch.

import { columnAxes } from '../cartesian.ts'
import { autoDecimals, formatNumber } from '../format.ts'
import type { ChartContext, LegendItem } from '../frame.ts'
import { chartFrame, fmt, unitsNote, warn } from '../frame.ts'
import { lineHeight, textWidth } from '../measure.ts'
import { INK } from '../palette.ts'
import { allValues, prepareSeries } from '../series.ts'
import { centredBaseline, group, num, path, text } from '../svg.ts'
import { placeholder } from './empty.ts'
import { spread } from './line.ts'

export type AreaMode = 'stacked' | 'percent'

export interface AreaLayer {
  name: string
  colour: string
  /** Bottom and top of the band at each category. */
  bottoms: number[]
  tops: number[]
}

/** Stacks series into bands. Gaps count as zero and negative values are left out; the flags say whether either happened. */
export function areaLayers(series: Array<{ name: string; colour: string; values: Array<number | null> }>, count: number, mode: AreaMode): { layers: AreaLayer[]; gaps: boolean; negatives: boolean } {
  let gaps = false
  let negatives = false
  const clean = series.map((s) => Array.from({ length: count }, (_, c) => {
    const v = s.values[c] ?? null
    if (v === null) gaps = true
    else if (v < 0) negatives = true
    return v === null || v < 0 ? 0 : v
  }))
  const totals = Array.from({ length: count }, (_, c) => clean.reduce((sum, values) => sum + (values[c] ?? 0), 0))
  const running = Array.from({ length: count }, () => 0)
  const layers = series.map((s, k) => {
    const bottoms = [...running]
    const tops = running.map((base, c) => {
      const raw = clean[k]?.[c] ?? 0
      const value = mode === 'percent' ? ((totals[c] ?? 0) > 0 ? (raw / (totals[c] ?? 1)) * 100 : 0) : raw
      return base + value
    })
    tops.forEach((t, c) => (running[c] = t))
    return { name: s.name, colour: s.colour, bottoms, tops }
  })
  return { layers, gaps, negatives }
}

export function renderArea(ctx: ChartContext, mode: AreaMode): string {
  const p = prepareSeries(ctx, { sortable: false, highlight: false })
  if (!p.series.length || !allValues(p.series).length) return placeholder(ctx, 'Add numbers to the data to draw the chart.')
  const n = p.categories.length
  if (n < 2) return placeholder(ctx, 'An area chart needs at least two categories, such as two years.')
  const { layers, gaps, negatives } = areaLayers(p.series, n, mode)
  if (gaps) warn(ctx, 'Empty cells are drawn as zero in an area chart.')
  if (negatives) warn(ctx, 'Area charts leave out negative values.')
  const font = ctx.font
  const decimals = ctx.format.decimals ?? autoDecimals(allValues(p.series))
  const bare = Boolean(unitsNote(ctx.format)) && mode === 'stacked'
  const items: LegendItem[] = layers.map((l) => ({ label: l.name, colour: l.colour }))
  const frame = chartFrame(ctx, items, bare ? unitsNote(ctx.format) : '')
  const last = n - 1
  const ends = layers
    .map((l, k) => ({ k, value: (l.tops[last] ?? 0) - (l.bottoms[last] ?? 0), mid: ((l.tops[last] ?? 0) + (l.bottoms[last] ?? 0)) / 2 }))
    .filter((e) => e.value > 0)
  const endText = (value: number): string =>
    mode === 'percent' ? formatNumber(value, { percent: true }, { decimals: 0 }) : fmt(ctx, value, { decimals, bare })
  const endRoom = ctx.labels && ends.length ? Math.max(...ends.map((e) => textWidth(endText(e.value), font))) + font * 0.6 : 0
  const hi = mode === 'percent' ? 100 : Math.max(0, ...layers.map((l) => Math.max(...l.tops)))
  const axes = columnAxes(ctx, { ...frame.plot, w: frame.plot.w - endRoom }, p.categories, { lo: 0, hi, percent: mode === 'percent' })
  const x = (i: number): number => axes.bandStart(i) + axes.band / 2
  const bands: string[] = []
  const seams: string[] = []
  layers.forEach((l, k) => {
    const upper = l.tops.map((t, i) => `${num(x(i))} ${num(axes.at(t))}`)
    const lower = l.bottoms.map((b, i) => `${num(x(i))} ${num(axes.at(b))}`).reverse()
    bands.push(path(`M${[...upper, ...lower].join('L')}Z`, l.colour))
    if (k < layers.length - 1) seams.push(`<path d="M${upper.join('L')}" fill="none" stroke="#${INK.white}" stroke-width="0.75" stroke-linejoin="round"/>`)
  })
  const labels: string[] = []
  if (ctx.labels && ends.length) {
    const ys = spread(ends.map((e) => axes.at(e.mid)), lineHeight(font) * 0.95, axes.plot.y, axes.plot.y + axes.plot.h)
    ends.forEach((e, i) => {
      labels.push(text(x(last) + font * 0.6, centredBaseline(ys[i] ?? axes.at(e.mid), font), endText(e.value), { size: font, fill: INK.data }))
    })
  }
  return frame.svg + axes.back + group('areas', bands.join('') + seams.join('')) + axes.front + group('labels', labels.join(''))
}
