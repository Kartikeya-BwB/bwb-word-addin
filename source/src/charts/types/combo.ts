// Combination charts: one or more series as columns and one or two as lines. The lines share the columns' axis, or
// with secondAxis read against an axis of their own on the right. With two axes, each carries a title above it and
// the legend marks the series that read against the right-hand one. Line labels are placed by the rules in place.ts.

import type { LinearScale } from '../axes.ts'
import { nearZero, niceScale, project, stepDecimals } from '../axes.ts'
import { columnAxes } from '../cartesian.ts'
import { columnValues } from '../data.ts'
import type { NumberFormat } from '../format.ts'
import { autoDecimals, formatNumber, unitsLabel } from '../format.ts'
import type { ChartContext, LegendItem } from '../frame.ts'
import { chartFrame, dropLabels, unitsNote, warn } from '../frame.ts'
import { MIN_FONT_PT, ellipsise, lineHeight, textWidth } from '../measure.ts'
import { INK } from '../palette.ts'
import type { Rect } from '../place.ts'
import { LabelPlacer, circleRect, textRect } from '../place.ts'
import type { DrawSeries } from '../series.ts'
import { prepareSeries } from '../series.ts'
import { circle, group, num, rect, text } from '../svg.ts'
import { maxBar } from './bars.ts'
import { placeholder } from './empty.ts'
import { lineStyle } from './line.ts'

/** The data series drawn as lines: the ones chosen (two at most, leaving at least one column), or else the last. */
export function lineIndices(seriesCount: number, chosen: number[] | undefined): number[] {
  if (seriesCount < 2) return []
  if (chosen === undefined) return [seriesCount - 1]
  const valid = [...new Set(chosen.filter((i) => Number.isInteger(i) && i >= 0 && i < seriesCount))]
  return valid.slice(0, Math.min(2, seriesCount - 1))
}

/** Units several series agree on, from their own cells. */
function sharedHint(grid: string[][], series: DrawSeries[]): NumberFormat {
  const hints = series.map((s) => columnValues(grid, s.index + 1).hint)
  const first = hints[0] ?? {}
  return hints.every((h) => JSON.stringify(h) === JSON.stringify(first)) ? first : {}
}

const tickText = (value: number, scale: LinearScale, format: NumberFormat): string =>
  formatNumber(value, { ...(format.percent ? { percent: true } : {}) }, { decimals: stepDecimals(scale.step) })

export function renderCombo(ctx: ChartContext): string {
  const p = prepareSeries(ctx, { sortable: false, highlight: false })
  const grid = ctx.spec.data
  const values = new Map(p.series.map((s) => [s.index, columnValues(grid, s.index + 1).values]))
  const numbers = [...values.values()].flat().filter((v): v is number => v !== null)
  if (!p.series.length || !numbers.length) return placeholder(ctx, 'Add numbers to the data to draw the chart.')
  const lineSet = new Set(lineIndices(p.table.series.length, ctx.spec.lineSeries))
  const columns = p.series.filter((s) => !lineSet.has(s.index))
  const lines = p.series.filter((s) => lineSet.has(s.index))
  if (!lines.length) warn(ctx, 'Choose a series to draw as a line; the chart shows columns only.')
  const font = ctx.font
  const lh = lineHeight(font)
  const second = Boolean(ctx.spec.secondAxis) && lines.length > 0
  const leftFormat: NumberFormat = { ...sharedHint(grid, columns), ...ctx.spec.format }
  const rightFormat: NumberFormat = second ? { ...sharedHint(grid, lines), ...ctx.spec.format2 } : leftFormat
  const local: ChartContext = { ...ctx, valueAxis: second || ctx.valueAxis, format: leftFormat }
  const leftNote = unitsNote(leftFormat)
  const bare = Boolean(leftNote) && local.valueAxis
  const items: LegendItem[] = [
    ...columns.map((s) => ({ label: s.name, colour: s.colour })),
    ...lines.map((s) => ({ label: second ? `${s.name} (right axis)` : s.name, colour: lineStyle(s).colour, mark: lineStyle(s).dashed ? ('dashed' as const) : ('marker' as const) })),
  ]
  const frame = chartFrame(local, items, second ? '' : bare ? leftNote : '')
  let area = frame.plot
  const titles: string[] = []
  if (second) {
    const leftTitle = ctx.spec.axisTitles?.y ?? (leftNote || columns.map((s) => s.name).join(', '))
    const rightUnits = unitsLabel(rightFormat)
    const rightTitle = ctx.spec.axisTitles?.y2 ?? `${lines.map((s) => s.name).join(', ')}${rightUnits ? ` (${rightUnits})` : ''}`
    const baseline = area.y + font * 0.86
    titles.push(text(area.x, baseline, ellipsise(leftTitle, area.w * 0.45, font, true), { size: font, fill: INK.text, bold: true }))
    titles.push(text(area.x + area.w, baseline, ellipsise(rightTitle, area.w * 0.45, font, true), { size: font, fill: INK.text, bold: true, anchor: 'end' }))
    area = { ...area, y: area.y + lh + font * 0.4, h: area.h - lh - font * 0.4 }
  }
  const of = (list: DrawSeries[]): number[] => list.flatMap((s) => (values.get(s.index) ?? []).filter((v): v is number => v !== null))
  const lineValues = of(lines)
  const leftValues = second ? of(columns) : numbers
  const rightScaleFor = (length: number): LinearScale => {
    const lo = Math.min(...lineValues)
    const hi = Math.max(...lineValues)
    return niceScale(lo, hi, { maxTicks: Math.min(8, Math.max(2, Math.floor(length / (font * 2.6)) + 1)), includeZero: nearZero(lo, hi) })
  }
  const rightWidth = second
    ? Math.max(...[area.h * 0.6, area.h].flatMap((h) => {
      const s = rightScaleFor(h)
      return s.ticks.map((t) => textWidth(tickText(t, s, rightFormat), font))
    })) + font * 0.6
    : 0
  const axes = columnAxes(local, { ...area, w: area.w - rightWidth }, p.categories, {
    lo: Math.min(0, ...leftValues), hi: Math.max(0, ...leftValues), roomAbove: ctx.labels,
  })
  const right = second ? rightScaleFor(axes.plot.h) : null
  const yRight = (v: number): number => (right ? project(right, Math.min(right.max, Math.max(right.min, v)), axes.plot.y + axes.plot.h, axes.plot.y) : axes.at(v))
  const rightTicks = right
    ? right.ticks.map((t) => text(axes.plot.x + axes.plot.w + font * 0.5, yRight(t) + font * 0.345, tickText(t, right, rightFormat), { size: font, fill: INK.text }))
    : []
  // Columns, clustered.
  const k = Math.max(1, columns.length)
  const groupWidth = Math.min(axes.band / (1 + ctx.gap), maxBar(ctx, k))
  const inner = k > 1 ? Math.min(1.5, groupWidth * 0.04) : 0
  const barWidth = (groupWidth - inner * (k - 1)) / k
  const decimals = leftFormat.decimals ?? autoDecimals(of(columns))
  const write = (v: number): string => formatNumber(v, leftFormat, { decimals, bare })
  const columnTexts = columns.flatMap((s) => (values.get(s.index) ?? []).filter((v): v is number => v !== null).map(write))
  let size: number | null = null
  for (let t = font; t >= MIN_FONT_PT && size === null; t -= 0.5) if (columnTexts.every((tx) => textWidth(tx, t) <= barWidth + inner - t * 0.2)) size = t
  if (ctx.labels && columnTexts.length && size === null) dropLabels(ctx)
  const bars: string[] = []
  const labels: string[] = []
  const obstacles: Rect[] = []
  p.categories.forEach((_, c) => {
    const x0 = axes.bandStart(c) + (axes.band - groupWidth) / 2
    columns.forEach((s, j) => {
      const v = values.get(s.index)?.[c] ?? null
      if (v === null) return
      const x = x0 + j * (barWidth + inner)
      const [ya, yb] = [axes.at(0 < axes.scale.min ? axes.scale.min : 0), axes.at(v)]
      bars.push(rect(x, Math.min(ya, yb), barWidth, Math.abs(yb - ya), s.colour))
      obstacles.push({ x0: x, y0: Math.min(ya, yb), x1: x + barWidth, y1: Math.max(ya, yb) })
      if (!ctx.labels || size === null) return
      const y = v >= 0 ? yb - size * 0.35 : yb + size * 1.04
      labels.push(text(x + barWidth / 2, y, write(v), { size, fill: INK.data, anchor: 'middle' }))
      obstacles.push(textRect(x + barWidth / 2, y, write(v), size, 'middle'))
    })
  })
  // Lines, with markers and labels placed where they are free.
  const centre = (i: number): number => axes.bandStart(i) + axes.band / 2
  const stroke = Math.max(1.5, font * 0.19)
  const markR = Math.max(2, font * 0.27)
  const drawn: string[] = []
  const lineFormat = second ? rightFormat : leftFormat
  const lineDecimals = lineFormat.decimals ?? autoDecimals(lineValues)
  const placer = new LabelPlacer({ x0: axes.plot.x, y0: axes.plot.y - lh, x1: axes.plot.x + axes.plot.w, y1: axes.plot.y + axes.plot.h }, obstacles)
  for (const s of lines) {
    const style = lineStyle(s)
    const points = (values.get(s.index) ?? []).map((v, i) => (v === null ? null : { x: centre(i), y: yRight(v), v }))
    const runs: string[][] = [[]]
    for (const pt of points) {
      if (pt === null) runs.push([])
      else runs[runs.length - 1]?.push(`${num(pt.x)} ${num(pt.y)}`)
    }
    const dash = style.dashed ? ` stroke-dasharray="${num(stroke * 3)} ${num(stroke * 2)}"` : ''
    for (const run of runs.filter((r) => r.length > 1)) {
      drawn.push(`<path d="M${run.join('L')}" fill="none" stroke="#${style.colour}" stroke-width="${num(stroke)}" stroke-linejoin="round" stroke-linecap="round"${dash}/>`)
    }
    for (const pt of points) if (pt) {
      drawn.push(circle(pt.x, pt.y, markR, style.colour, { colour: INK.white, width: Math.max(0.75, font * 0.08) }))
      placer.block(circleRect(pt.x, pt.y, markR))
    }
  }
  if (ctx.labels) {
    for (const s of lines) {
      for (const [i, v] of (values.get(s.index) ?? []).entries()) {
        if (v === null) continue
        const label = formatNumber(v, lineFormat, { decimals: lineDecimals, bare: !second && bare })
        const placed = placer.place(centre(i), yRight(v), markR, label, font, { sides: ['above', 'below', 'right', 'left'] })
        if (placed) labels.push(text(placed.x, placed.y, label, { size: font, fill: INK.data, anchor: placed.anchor }))
      }
    }
    if (placer.dropped) warn(ctx, `${placer.dropped} line label${placer.dropped === 1 ? ' is' : 's are'} left out where ${placer.dropped === 1 ? 'it' : 'they'} would overlap.`)
  }
  return frame.svg + group('titles', titles.join('')) + axes.back + group('bars', bars.join('')) + axes.front +
    group('right-axis', rightTicks.join('')) + group('lines', drawn.join('')) + group('labels', labels.join(''))
}
