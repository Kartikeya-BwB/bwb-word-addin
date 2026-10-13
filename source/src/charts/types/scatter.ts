// Scatter and bubble charts: one point per row, placed by two measures, with the row's name beside it. A bubble's
// area shows a third measure. Optional quadrant lines, at chosen values or at the medians, make the chart a 2x2 with
// a label in each corner. Names are placed by the rules in place.ts and left out, with a warning, when no side is free.

import { COLOURS } from '../../brand/tokens.ts'
import type { LinearScale } from '../axes.ts'
import { nearZero, niceScale, project, stepDecimals } from '../axes.ts'
import { columnValues } from '../data.ts'
import type { NumberFormat } from '../format.ts'
import { formatNumber } from '../format.ts'
import type { Box, ChartContext } from '../frame.ts'
import { chartFrame, warn } from '../frame.ts'
import type { Grid } from '../grid.ts'
import { normaliseGrid } from '../grid.ts'
import { ellipsise, lineHeight, textWidth } from '../measure.ts'
import { INK } from '../palette.ts'
import type { Rect } from '../place.ts'
import { LabelPlacer, circleRect, textRect } from '../place.ts'
import { centredBaseline, circle, group, line, text } from '../svg.ts'
import { placeholder } from './empty.ts'

export interface ScatterPoint {
  label: string
  x: number
  y: number
  /** Bubble measure; null on a plain scatter. */
  size: number | null
}

export interface ScatterData {
  points: ScatterPoint[]
  xTitle: string
  yTitle: string
  sizeTitle: string
  xHint: NumberFormat
  yHint: NumberFormat
  /** Rows left out because a measure was missing (or a bubble size was not above zero). */
  skipped: number
}

/** Reads points from a grid: names in the first column, then x, y and (for bubbles) size, headings in the first row. */
export function scatterData(input: Grid, bubble: boolean): ScatterData {
  const grid = normaliseGrid(input)
  const header = grid[0] ?? []
  const rows = grid.slice(1).filter((r) => r.some((c) => c.trim()))
  const [xs, ys, sizes] = [1, 2, 3].map((c) => columnValues(grid, c))
  const points: ScatterPoint[] = []
  rows.forEach((row, i) => {
    const x = xs?.values[i] ?? null
    const y = ys?.values[i] ?? null
    const size = bubble ? (sizes?.values[i] ?? null) : null
    if (x === null || y === null || (bubble && (size === null || size <= 0))) return
    points.push({ label: (row[0] ?? '').trim() || `Point ${i + 1}`, x, y, size })
  })
  return {
    points,
    xTitle: (header[1] ?? '').trim(),
    yTitle: (header[2] ?? '').trim(),
    sizeTitle: (header[3] ?? '').trim(),
    xHint: xs?.hint ?? {},
    yHint: ys?.hint ?? {},
    skipped: rows.length - points.length,
  }
}

export function median(values: number[]): number {
  const sorted = [...values].sort((a, b) => a - b)
  const mid = Math.floor(sorted.length / 2)
  return sorted.length % 2 ? (sorted[mid] ?? 0) : ((sorted[mid - 1] ?? 0) + (sorted[mid] ?? 0)) / 2
}

/** Tick labels carry the scale and per cent, not the currency, which belongs in the axis title. */
const tick = (value: number, scale: LinearScale, hint: NumberFormat): string =>
  formatNumber(value, { ...(hint.scale ? { scale: hint.scale } : {}), ...(hint.percent ? { percent: true } : {}) }, { decimals: stepDecimals(scale.step) })

/** A scale for one axis, widened by `margin` points at each end so bubbles at the extremes stay inside. */
function axisScale(lo: number, hi: number, length: number, margin: number, perTick: number): LinearScale {
  const maxTicks = Math.min(8, Math.max(2, Math.floor(length / perTick) + 1))
  let scale = niceScale(lo, hi, { maxTicks, includeZero: nearZero(lo, hi) })
  const pad = ((scale.max - scale.min) * margin) / Math.max(1, length - 2 * margin)
  if (pad > 0) scale = niceScale(lo - pad, hi + pad, { maxTicks, includeZero: nearZero(lo, hi) })
  return scale
}

export function renderScatter(ctx: ChartContext, bubble: boolean): string {
  const data = scatterData(ctx.spec.data, bubble)
  const what = bubble ? 'an x value, a y value and a size above zero' : 'an x value and a y value'
  if (data.skipped) warn(ctx, `${data.skipped} row${data.skipped === 1 ? ' is' : 's are'} left out: each point needs ${what}.`)
  if (!data.points.length) return placeholder(ctx, `Add rows with a name, ${what}.`)
  const font = ctx.font
  const lh = lineHeight(font)
  const showTitles = ctx.spec.showAxisTitles ?? true
  const xTitle = showTitles ? (ctx.spec.axisTitles?.x ?? data.xTitle) : ''
  const yTitle = showTitles ? (ctx.spec.axisTitles?.y ?? data.yTitle) : ''
  const note = bubble && data.sizeTitle ? `Bubble area shows ${data.sizeTitle}` : ''
  const frame = chartFrame(ctx, [], note)
  const area = frame.plot
  const top = area.y + (yTitle ? lh + font * 0.4 : font * 0.6)
  const bottom = area.y + area.h - lh - font * 0.3 - (xTitle ? lh + font * 0.2 : 0)
  const xs = data.points.map((p) => p.x)
  const ys = data.points.map((p) => p.y)
  const largest = Math.max(...data.points.map((p) => p.size ?? 0))
  const pointR = Math.max(3, font * 0.36)
  const rMax = bubble ? Math.max(font, Math.min(area.w, bottom - top) * 0.085) : pointR
  let yScale = axisScale(Math.min(...ys), Math.max(...ys), bottom - top, rMax, font * 2.6)
  let axisW = 0
  let plot: Box = { x: area.x, y: top, w: area.w, h: bottom - top }
  let xScale = axisScale(Math.min(...xs), Math.max(...xs), plot.w, rMax, font * 6)
  for (let pass = 0; pass < 2; pass += 1) {
    axisW = Math.max(...yScale.ticks.map((t) => textWidth(tick(t, yScale, data.yHint), font)), 0) + font * 0.6
    const lastTick = xScale.ticks[xScale.ticks.length - 1]
    const right = lastTick === undefined ? 0 : textWidth(tick(lastTick, xScale, data.xHint), font) / 2
    plot = { x: area.x + axisW, y: top, w: Math.max(font * 4, area.w - axisW - right), h: Math.max(font * 3, bottom - top) }
    const widest = Math.max(...xScale.ticks.map((t) => textWidth(tick(t, xScale, data.xHint), font)), font)
    xScale = axisScale(Math.min(...xs), Math.max(...xs), plot.w, rMax, widest + font * 2)
    yScale = axisScale(Math.min(...ys), Math.max(...ys), plot.h, rMax, font * 2.6)
  }
  const xAt = (v: number): number => project(xScale, v, plot.x, plot.x + plot.w)
  const yAt = (v: number): number => project(yScale, v, plot.y + plot.h, plot.y)
  const back: string[] = []
  const front: string[] = []
  for (const t of yScale.ticks) {
    const y = yAt(t)
    if (t !== yScale.min) back.push(line(plot.x, y, plot.x + plot.w, y, t === 0 ? INK.charcoal : INK.grid, t === 0 ? 0.75 : 0.5))
    front.push(text(plot.x - font * 0.5, centredBaseline(y, font), tick(t, yScale, data.yHint), { size: font, fill: INK.text, anchor: 'end' }))
  }
  back.push(line(plot.x, plot.y + plot.h, plot.x + plot.w, plot.y + plot.h, INK.charcoal, 0.75))
  if (xScale.min < 0 && xScale.max > 0) back.push(line(xAt(0), plot.y, xAt(0), plot.y + plot.h, INK.charcoal, 0.75))
  for (const t of xScale.ticks) {
    front.push(text(xAt(t), plot.y + plot.h + font * 0.3 + font * 0.86, tick(t, xScale, data.xHint), { size: font, fill: INK.text, anchor: 'middle' }))
  }
  if (yTitle) front.push(text(area.x, area.y + font * 0.86, ellipsise(yTitle, area.w * 0.7, font), { size: font, fill: INK.text, bold: true }))
  if (xTitle) front.push(text(plot.x + plot.w / 2, area.y + area.h - font * 0.22, ellipsise(xTitle, plot.w, font), { size: font, fill: INK.text, anchor: 'middle', bold: true }))
  const obstacles: Rect[] = []
  // Quadrant lines and corner labels.
  const q = ctx.spec.quadrants
  const qx = q?.x === 'median' ? median(xs) : q?.x
  const qy = q?.y === 'median' ? median(ys) : q?.y
  if (qx !== undefined && qx >= xScale.min && qx <= xScale.max) back.push(line(xAt(qx), plot.y, xAt(qx), plot.y + plot.h, INK.charcoal, 0.75, '3 2'))
  if (qy !== undefined && qy >= yScale.min && qy <= yScale.max) back.push(line(plot.x, yAt(qy), plot.x + plot.w, yAt(qy), INK.charcoal, 0.75, '3 2'))
  const corner = Math.max(8, font * 0.9)
  const pad = font * 0.4
  const corners: Array<{ x: number; y: number; anchor: 'start' | 'end' }> = [
    { x: plot.x + pad, y: plot.y + pad + corner * 0.72, anchor: 'start' },
    { x: plot.x + plot.w - pad, y: plot.y + pad + corner * 0.72, anchor: 'end' },
    { x: plot.x + pad, y: plot.y + plot.h - pad - corner * 0.22, anchor: 'start' },
    { x: plot.x + plot.w - pad, y: plot.y + plot.h - pad - corner * 0.22, anchor: 'end' },
  ]
  q?.labels?.slice(0, 4).forEach((raw, i) => {
    const spot = corners[i]
    const label = ellipsise(raw.trim(), plot.w * 0.45, corner, true)
    if (!spot || !label) return
    front.push(text(spot.x, spot.y, label, { size: corner, fill: INK.charcoal, anchor: spot.anchor, bold: true }))
    obstacles.push(textRect(spot.x, spot.y, label, corner, spot.anchor, true))
  })
  // Points, largest bubbles first so smaller ones stay visible on top.
  const wanted = ctx.spec.highlight?.trim().toLowerCase() ?? ''
  const order = [...data.points].sort((a, b) => (b.size ?? 0) - (a.size ?? 0))
  const radius = (p: ScatterPoint): number => (bubble && largest > 0 ? Math.max(1.5, rMax * Math.sqrt((p.size ?? 0) / largest)) : pointR)
  const marks = order.map((p) => {
    const highlight = wanted !== '' && p.label.toLowerCase() === wanted
    const [cx, cy, r] = [xAt(p.x), yAt(p.y), radius(p)]
    obstacles.push(circleRect(cx, cy, r))
    return { p, cx, cy, r, highlight, svg: circle(cx, cy, r, highlight ? COLOURS.green : COLOURS.navy, { colour: INK.white, width: bubble ? 1 : 0.75 }) }
  })
  const labels: string[] = []
  if (ctx.labels) {
    const placer = new LabelPlacer({ x0: plot.x, y0: plot.y, x1: plot.x + plot.w, y1: plot.y + plot.h }, obstacles)
    const byImportance = [...marks].sort((a, b) => Number(b.highlight) - Number(a.highlight))
    for (const m of byImportance) {
      const label = ellipsise(m.p.label, plot.w * 0.4, font, m.highlight)
      const placed = placer.place(m.cx, m.cy, m.r, label, font, { bold: m.highlight })
      if (placed) labels.push(text(placed.x, placed.y, label, { size: font, fill: INK.data, anchor: placed.anchor, bold: m.highlight }))
    }
    if (placer.dropped) {
      warn(ctx, `${placer.dropped} point label${placer.dropped === 1 ? ' is' : 's are'} left out where ${placer.dropped === 1 ? 'it' : 'they'} would overlap. Make the chart bigger or the names shorter.`)
    }
  }
  return frame.svg + group('gridlines', back.join('')) + group('points', marks.map((m) => m.svg).join('')) + group('axes', front.join('')) + group('labels', labels.join(''))
}
