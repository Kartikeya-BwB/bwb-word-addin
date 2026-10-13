// Line charts: categories along the bottom, one line per series, gaps where a value is missing. With one series every
// point is labelled, above the line or below it in a dip; with several, each line is labelled at its last point and
// the labels are spread so they never touch.

import { COLOURS } from '../../brand/tokens.ts'
import { columnAxes } from '../cartesian.ts'
import { autoDecimals } from '../format.ts'
import type { ChartContext, LegendItem } from '../frame.ts'
import { chartFrame, dropLabels, fmt, unitsNote } from '../frame.ts'
import { CAP_HEIGHT, MIN_FONT_PT, lineHeight, textWidth } from '../measure.ts'
import { INK } from '../palette.ts'
import type { DrawSeries } from '../series.ts'
import { allValues, prepareSeries } from '../series.ts'
import { centredBaseline, circle, group, num, text } from '../svg.ts'
import { placeholder } from './empty.ts'

interface LineStyle {
  colour: string
  dashed: boolean
}

/** Light grey and charcoal 20 lines would vanish on white, so those series are drawn as dashed navy and green. */
export function lineStyle(series: DrawSeries): LineStyle {
  if (series.colour === COLOURS.lightGrey) return { colour: COLOURS.navy, dashed: true }
  if (series.colour === COLOURS.charcoal20) return { colour: COLOURS.green, dashed: true }
  return { colour: series.colour, dashed: false }
}

/** Zero stays on the axis unless the values sit far from it, where it would flatten the line. */
function keepZero(lo: number, hi: number): boolean {
  if (lo >= 0) return lo <= hi * 0.5
  if (hi <= 0) return hi >= lo * 0.5
  return true
}

/** Moves labels apart so that each is at least `gap` below the one above, keeping them inside top to bottom. */
export function spread(ys: number[], gap: number, top: number, bottom: number): number[] {
  const order = ys.map((y, i) => ({ y, i })).sort((a, b) => a.y - b.y)
  for (let k = 1; k < order.length; k += 1) {
    const prev = order[k - 1]
    const cur = order[k]
    if (prev && cur && cur.y - prev.y < gap) cur.y = prev.y + gap
  }
  const last = order[order.length - 1]
  const overflow = last ? last.y - bottom : 0
  if (overflow > 0) for (const o of order) o.y -= overflow
  const first = order[0]
  if (first && first.y < top) {
    const shift = top - first.y
    for (const o of order) o.y += shift
  }
  const out = [...ys]
  for (const o of order) out[o.i] = o.y
  return out
}

export function renderLine(ctx: ChartContext): string {
  const p = prepareSeries(ctx, { sortable: false, highlight: false })
  const values = allValues(p.series)
  if (!values.length) return placeholder(ctx, 'Add numbers to the data to draw the chart.')
  const font = ctx.font
  const markers = ctx.spec.markers ?? true
  const styles = p.series.map(lineStyle)
  const many = p.series.length > 1
  const decimals = ctx.format.decimals ?? autoDecimals(values)
  const bare = Boolean(unitsNote(ctx.format)) && (ctx.valueAxis || many)
  const write = (v: number): string => fmt(ctx, v, { decimals, bare })
  const items: LegendItem[] = p.series.map((s, i) => ({
    label: s.name,
    colour: styles[i]?.colour ?? s.colour,
    mark: styles[i]?.dashed ? 'dashed' : markers ? 'marker' : 'line',
  }))
  const frame = chartFrame(ctx, items, bare ? unitsNote(ctx.format) : '')
  const lastIndex = (s: DrawSeries): number => {
    for (let i = s.values.length - 1; i >= 0; i -= 1) if (s.values[i] !== null && s.values[i] !== undefined) return i
    return -1
  }
  const ended = many && ctx.labels ? p.series.filter((s) => lastIndex(s) >= 0) : []
  const endTexts = ended.map((s) => write(s.values[lastIndex(s)] ?? 0))
  const endRoom = endTexts.length ? Math.max(...endTexts.map((t) => textWidth(t, font))) + font * 0.6 : 0
  const lo = Math.min(...values)
  const hi = Math.max(...values)
  const area = { ...frame.plot, w: frame.plot.w - endRoom }
  const axes = columnAxes(ctx, area, p.categories, {
    lo, hi, includeZero: keepZero(lo, hi), roomAbove: ctx.labels && !many, roomBelow: ctx.labels && !many,
  })
  const x = (i: number): number => axes.bandStart(i) + axes.band / 2
  const stroke = Math.max(1.5, font * 0.19)
  const lines: string[] = []
  const dots: string[] = []
  p.series.forEach((s, k) => {
    const style = styles[k] ?? { colour: s.colour, dashed: false }
    const runs: string[][] = [[]]
    s.values.forEach((v, i) => {
      if (v === null || v === undefined) runs.push([])
      else runs[runs.length - 1]?.push(`${num(x(i))} ${num(axes.at(v))}`)
    })
    const dash = style.dashed ? ` stroke-dasharray="${num(stroke * 3)} ${num(stroke * 2)}"` : ''
    for (const run of runs) {
      if (run.length > 1) {
        lines.push(`<path d="M${run.join('L')}" fill="none" stroke="#${style.colour}" stroke-width="${num(stroke)}" stroke-linejoin="round" stroke-linecap="round"${dash}/>`)
      }
      if (run.length === 1 && !markers) {
        const [px, py] = (run[0] ?? '0 0').split(' ').map(Number)
        dots.push(circle(px ?? 0, py ?? 0, stroke * 1.2, style.colour))
      }
    }
    if (markers) {
      s.values.forEach((v, i) => {
        if (v !== null && v !== undefined) dots.push(circle(x(i), axes.at(v), Math.max(2, font * 0.27), style.colour, { colour: INK.white, width: Math.max(0.75, font * 0.08) }))
      })
    }
  })
  const labels: string[] = []
  if (ctx.labels && !many) {
    const s = p.series[0]
    const texts = (s?.values ?? []).map((v) => (v === null ? '' : write(v)))
    let size: number | null = null
    for (let t = font; t >= MIN_FONT_PT && size === null; t -= 0.5) if (texts.every((tx) => textWidth(tx, t) <= axes.band - t * 0.3)) size = t
    if (size === null) dropLabels(ctx)
    s?.values.forEach((v, i) => {
      if (v === null || size === null) return
      const before = s.values[i - 1] ?? null
      const after = s.values[i + 1] ?? null
      const dip = (before === null || before > v) && (after === null || after > v) && (before !== null || after !== null)
      const offset = Math.max(2, font * 0.27) + size * 0.35
      const y = dip ? axes.at(v) + offset + size * CAP_HEIGHT : axes.at(v) - offset
      labels.push(text(x(i), y, write(v), { size, fill: INK.data, anchor: 'middle' }))
    })
  }
  if (endTexts.length) {
    const ends = ended.map((s) => {
      const i = lastIndex(s)
      return { x: x(i), y: axes.at(s.values[i] ?? 0) }
    })
    const ys = spread(ends.map((e) => e.y), lineHeight(font) * 0.95, axes.plot.y, axes.plot.y + axes.plot.h)
    ends.forEach((e, k) => {
      const at = Math.max(...ends.map((d) => d.x)) + font * 0.6
      labels.push(text(at, centredBaseline(ys[k] ?? e.y, font), endTexts[k] ?? '', { size: font, fill: INK.data }))
    })
  }
  return frame.svg + axes.back + group('lines', lines.join('') + dots.join('')) + axes.front + group('labels', labels.join(''))
}
