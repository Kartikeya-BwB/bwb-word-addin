// Column and bar charts: clustered, stacked and 100% stacked. Data labels sit outside single bars and inside stack
// segments (white on dark fills, navy on light ones), with stack totals beyond the stack. A label that cannot fit
// at 8 pt or more is left out rather than allowed to collide.

import { barAxes, columnAxes } from '../cartesian.ts'
import { autoDecimals, formatNumber } from '../format.ts'
import type { ChartContext, LegendItem } from '../frame.ts'
import { chartFrame, dropLabels, fmt, unitsNote, warn } from '../frame.ts'
import { CAP_HEIGHT, MIN_FONT_PT, lineHeight, textWidth } from '../measure.ts'
import { INK, inkOn } from '../palette.ts'
import type { DrawSeries, Prepared } from '../series.ts'
import { allValues, prepareSeries } from '../series.ts'
import { centredBaseline, group, line, rect, text } from '../svg.ts'
import { placeholder } from './empty.ts'

export type BarMode = 'clustered' | 'stacked' | 'percent'

interface Segment {
  series: DrawSeries
  /** Value drawn: the number, or its share for 100% charts. */
  value: number
  from: number
  to: number
  colour: string
}

interface Stack {
  segments: Segment[]
  total: number
  top: number
  bottom: number
}

function stacks(ctx: ChartContext, p: Prepared, mode: BarMode): Stack[] {
  let dropped = false
  const out = p.categories.map((_, c) => {
    const values = p.series.map((s) => s.values[c] ?? null)
    if (mode === 'percent' && values.some((v) => v !== null && v < 0)) dropped = true
    const whole = values.reduce<number>((sum, v) => sum + Math.max(0, v ?? 0), 0)
    let up = 0
    let down = 0
    const segments: Segment[] = []
    p.series.forEach((s, i) => {
      const raw = values[i]
      if (raw === null || raw === undefined || raw === 0) return
      const value = mode === 'percent' ? (whole > 0 ? (Math.max(0, raw) / whole) * 100 : 0) : raw
      if (value === 0) return
      if (mode === 'clustered') segments.push({ series: s, value, from: 0, to: value, colour: p.colourAt(s, c) })
      else if (value > 0) segments.push({ series: s, value, from: up, to: (up += value), colour: p.colourAt(s, c) })
      else segments.push({ series: s, value, from: down, to: (down += value), colour: p.colourAt(s, c) })
    })
    const total = values.reduce<number>((sum, v) => sum + (v ?? 0), 0)
    const ends = segments.map((g) => g.to)
    return { segments, total, top: Math.max(0, ...ends), bottom: Math.min(0, ...ends) }
  })
  if (dropped) warn(ctx, '100% charts leave out negative values.')
  return out
}

const legendItems = (p: Prepared): LegendItem[] => p.series.map((s) => ({ label: s.name, colour: s.colour }))

/** The largest label size from the chart's size down to 8 pt at which every text fits its width, or null. */
function fitSize(texts: string[], room: number, start: number): number | null {
  for (let size = start; size >= MIN_FONT_PT; size -= 0.5) {
    if (texts.every((t) => textWidth(t, size) <= room)) return size
  }
  return null
}

interface LabelPlan {
  /** Labels without currency and scale, with the units written once at the top left. */
  bare: boolean
  decimals: number
}

/** Leaves out the units from labels when the value axis shows numbers, or when labels with units would not fit. */
function planLabels(ctx: ChartContext, values: number[], room: number): LabelPlan {
  const decimals = ctx.format.decimals ?? autoDecimals(values)
  if (!unitsNote(ctx.format)) return { bare: false, decimals }
  if (ctx.valueAxis) return { bare: true, decimals }
  const full = values.map((v) => fmt(ctx, v, { decimals }))
  return { bare: fitSize(full, room, ctx.font) === null, decimals }
}

const shareText = (value: number): string => formatNumber(value, { percent: true }, { decimals: 0 })

/** Columns: categories along the bottom. */
export function renderColumns(ctx: ChartContext, mode: BarMode): string {
  const p = prepareSeries(ctx, { sortable: true, highlight: mode === 'clustered' })
  const all = stacks(ctx, p, mode)
  if (!p.series.length || !allValues(p.series).length) return placeholder(ctx, 'Add numbers to the data to draw the chart.')
  const n = p.categories.length
  const roughBand = (ctx.width * 0.9) / Math.max(1, n) / (1 + ctx.gap)
  const roughRoom = mode === 'clustered' && p.series.length > 1 ? roughBand / p.series.length : roughBand
  const plan = planLabels(ctx, mode === 'stacked' ? all.map((s) => s.total) : allValues(p.series), roughRoom)
  const write = (v: number): string => fmt(ctx, v, { decimals: plan.decimals, bare: plan.bare })
  const frame = chartFrame(ctx, legendItems(p), plan.bare ? unitsNote(ctx.format) : '')
  const lo = Math.min(0, ...all.map((s) => s.bottom))
  const hi = Math.max(0, ...all.map((s) => s.top))
  const outside = ctx.labels && mode !== 'percent'
  const axes = columnAxes(ctx, frame.plot, p.categories, {
    lo, hi, percent: mode === 'percent', roomAbove: outside, roomBelow: outside && lo < 0,
  })
  const clusters = mode === 'clustered' ? p.series.length : 1
  const groupWidth = Math.min(axes.band / (1 + ctx.gap), maxBar(ctx, clusters))
  const inner = clusters > 1 ? Math.min(1.5, groupWidth * 0.04) : 0
  const barWidth = (groupWidth - inner * (clusters - 1)) / clusters
  const bars: string[] = []
  const labels: string[] = []
  const outsideTexts: string[] = []
  all.forEach((stack) => {
    if (mode === 'stacked' && p.series.length > 1 && stack.segments.length) outsideTexts.push(write(stack.total))
    if (mode === 'clustered' || p.series.length === 1) stack.segments.forEach((g) => outsideTexts.push(write(g.value)))
  })
  const outsideRoom = clusters > 1 ? barWidth + inner : axes.band - ctx.font * 0.4
  const outsideSize = outside ? fitSize(outsideTexts, outsideRoom, ctx.font) : null
  if (outside && outsideTexts.length && outsideSize === null) dropLabels(ctx)
  all.forEach((stack, c) => {
    const x0 = axes.bandStart(c) + (axes.band - groupWidth) / 2
    stack.segments.forEach((g, k) => {
      const slot = mode === 'clustered' ? p.series.indexOf(g.series) : 0
      const x = x0 + slot * (barWidth + inner)
      const yA = axes.at(g.from)
      const yB = axes.at(g.to)
      bars.push(rect(x, Math.min(yA, yB), barWidth, Math.abs(yB - yA), g.colour))
      if (mode !== 'clustered' && k > 0) bars.push(line(x, yA, x + barWidth, yA, INK.white, 0.75))
      if (!ctx.labels) return
      const cx = x + barWidth / 2
      if (mode === 'clustered' || (mode === 'stacked' && p.series.length === 1)) {
        if (outsideSize === null) return
        const y = g.value >= 0 ? yB - outsideSize * 0.35 : yB + outsideSize * (0.35 + CAP_HEIGHT)
        labels.push(text(cx, y, write(g.value), { size: outsideSize, fill: INK.data, anchor: 'middle' }))
        return
      }
      const label = mode === 'percent' ? shareText(g.value) : write(g.value)
      const size = ctx.font
      if (textWidth(label, size) <= barWidth - size * 0.3 && Math.abs(yB - yA) >= size * (CAP_HEIGHT + 0.5)) {
        labels.push(text(cx, centredBaseline((yA + yB) / 2, size), label, { size, fill: inkOn(g.colour), anchor: 'middle' }))
      }
    })
    if (mode === 'stacked' && p.series.length > 1 && ctx.labels && outsideSize !== null && stack.segments.length) {
      // The total sits beyond the end the stack grows towards: above it, or below an all-negative stack.
      const down = stack.top === 0 && stack.bottom < 0
      const y = down ? axes.at(stack.bottom) + outsideSize * (0.35 + CAP_HEIGHT) : axes.at(stack.top) - outsideSize * 0.35
      labels.push(text(x0 + groupWidth / 2, y, write(stack.total), { size: outsideSize, fill: INK.data, anchor: 'middle', bold: true }))
    }
  })
  return frame.svg + axes.back + group('bars', bars.join('')) + axes.front + group('labels', labels.join(''))
}

/** Bars: categories down the left, the first at the top. */
export function renderBars(ctx: ChartContext, mode: BarMode): string {
  const p = prepareSeries(ctx, { sortable: true, highlight: mode === 'clustered' })
  const all = stacks(ctx, p, mode)
  if (!p.series.length || !allValues(p.series).length) return placeholder(ctx, 'Add numbers to the data to draw the chart.')
  const plan = planLabels(ctx, mode === 'stacked' ? all.map((s) => s.total) : allValues(p.series), Number.POSITIVE_INFINITY)
  const write = (v: number): string => fmt(ctx, v, { decimals: plan.decimals, bare: plan.bare })
  const frame = chartFrame(ctx, legendItems(p), plan.bare ? unitsNote(ctx.format) : '')
  const lo = Math.min(0, ...all.map((s) => s.bottom))
  const hi = Math.max(0, ...all.map((s) => s.top))
  const font = ctx.font
  const outside = ctx.labels && mode !== 'percent'
  const ends: Array<{ text: string; negative: boolean }> = []
  all.forEach((stack) => {
    if (mode === 'clustered' || p.series.length === 1) stack.segments.forEach((g) => ends.push({ text: write(g.value), negative: g.value < 0 }))
    else if (mode === 'stacked' && stack.segments.length) ends.push({ text: write(stack.total), negative: stack.total < 0 && stack.top === 0 })
  })
  const widest = (negative: boolean): number =>
    Math.max(0, ...ends.filter((e) => e.negative === negative).map((e) => textWidth(e.text, font, mode === 'stacked')))
  const axes = barAxes(ctx, frame.plot, p.categories, {
    lo, hi, percent: mode === 'percent',
    labelRoom: outside ? widest(false) + font * 0.5 : font * 0.5,
    negativeRoom: outside && lo < 0 ? widest(true) + font * 1.2 : 0,
  })
  const clusters = mode === 'clustered' ? p.series.length : 1
  const groupHeight = Math.min(axes.band / (1 + ctx.gap), maxBar(ctx, clusters) * 0.4)
  const inner = clusters > 1 ? Math.min(1.5, groupHeight * 0.04) : 0
  const barHeight = (groupHeight - inner * (clusters - 1)) / clusters
  const pitch = clusters > 1 ? barHeight + inner : axes.band
  const outsideSize = outside ? sizeForPitch(pitch, font) : null
  if (outside && ends.length && outsideSize === null) dropLabels(ctx)
  const bars: string[] = []
  const labels: string[] = []
  all.forEach((stack, c) => {
    const y0 = axes.bandStart(c) + (axes.band - groupHeight) / 2
    stack.segments.forEach((g, k) => {
      const slot = mode === 'clustered' ? p.series.indexOf(g.series) : 0
      const y = y0 + slot * (barHeight + inner)
      const xA = axes.at(g.from)
      const xB = axes.at(g.to)
      bars.push(rect(Math.min(xA, xB), y, Math.abs(xB - xA), barHeight, g.colour))
      if (mode !== 'clustered' && k > 0) bars.push(line(xA, y, xA, y + barHeight, INK.white, 0.75))
      if (!ctx.labels) return
      const cy = y + barHeight / 2
      if (mode === 'clustered' || (mode === 'stacked' && p.series.length === 1)) {
        if (outsideSize === null) return
        const gap = outsideSize * 0.4
        const anchor = g.value >= 0 ? 'start' : 'end'
        labels.push(text(xB + (g.value >= 0 ? gap : -gap), centredBaseline(cy, outsideSize), write(g.value), { size: outsideSize, fill: INK.data, anchor }))
        return
      }
      const label = mode === 'percent' ? shareText(g.value) : write(g.value)
      if (textWidth(label, font) <= Math.abs(xB - xA) - font * 0.4 && barHeight >= font * (CAP_HEIGHT + 0.4)) {
        labels.push(text((xA + xB) / 2, centredBaseline(cy, font), label, { size: font, fill: inkOn(g.colour), anchor: 'middle' }))
      }
    })
    if (mode === 'stacked' && p.series.length > 1 && outsideSize !== null && stack.segments.length) {
      const down = stack.top === 0 && stack.bottom < 0
      const x = down ? axes.at(stack.bottom) - outsideSize * 0.4 : axes.at(stack.top) + outsideSize * 0.4
      labels.push(text(x, centredBaseline(y0 + groupHeight / 2, outsideSize), write(stack.total), { size: outsideSize, fill: INK.data, bold: true, anchor: down ? 'end' : 'start' }))
    }
  })
  return frame.svg + axes.back + group('bars', bars.join('')) + axes.front + group('labels', labels.join(''))
}

/** The widest a bar (or a cluster of bars) gets, so a chart of a few categories does not turn into blocks. */
export function maxBar(ctx: ChartContext, clusters: number): number {
  return ctx.font * (clusters > 1 ? 3.5 * clusters + 1 : 7)
}

/** Label size for labels stacked a pitch apart, or null when even 8 pt labels would touch. */
function sizeForPitch(pitch: number, font: number): number | null {
  for (let size = font; size >= MIN_FONT_PT; size -= 0.5) if (lineHeight(size) * 0.85 <= pitch) return size
  return null
}
