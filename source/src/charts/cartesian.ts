// Category and value axes shared by column, line and waterfall charts (categories along the bottom) and bar charts
// (categories down the left). Category labels wrap to their band, so they never run into each other; gridlines are
// light grey and horizontal only, so bar charts have none.

import type { LinearScale } from './axes.ts'
import { niceScale, project, stepDecimals } from './axes.ts'
import { formatNumber } from './format.ts'
import { abbreviateMonths } from './dates.ts'
import type { Box, ChartContext } from './frame.ts'
import { LABELS_CUT, warn } from './frame.ts'
import { fitLabels, lineHeight, textWidth } from './measure.ts'
import { INK } from './palette.ts'
import { centredBaseline, group, line, text, textLines, textLinesCentred } from './svg.ts'

export interface AxisOptions {
  lo: number
  hi: number
  /** Keep zero on the axis. Default true; line charts far from zero leave it out. */
  includeZero?: boolean
  /** A 0 to 100% axis for 100% charts. */
  percent?: boolean
  /** Room for data labels beyond the largest values. */
  roomAbove?: boolean
  /** Room for data labels beyond negative values. */
  roomBelow?: boolean
  /** Extra room above the labels, in points, for annotation arrows. */
  extraTop?: number
}

export interface Axes {
  plot: Box
  scale: LinearScale
  /** Coordinate of a value along the value axis. */
  at: (value: number) => number
  /** Coordinate of zero, or of the axis minimum when zero is off the axis. */
  base: number
  band: number
  /** Start of a category's band along the category axis. */
  bandStart: (index: number) => number
  /** Gridlines, drawn behind the data. */
  back: string
  /** Tick labels, the zero line and category labels, drawn over the data. */
  front: string
  /** Where the room for data labels beyond the data ends: the top of it for columns, the right for bars. */
  edge: number
}

function tickText(ctx: ChartContext, value: number, step: number, percent: boolean): string {
  return formatNumber(value, { percent: percent || Boolean(ctx.format.percent) }, { decimals: stepDecimals(step) })
}

function makeScale(ctx: ChartContext, o: AxisOptions, length: number, perTick: number): LinearScale {
  const maxTicks = Math.min(8, Math.max(2, Math.floor(length / perTick) + 1))
  if (o.percent) return niceScale(0, 100, { min: 0, max: 100, maxTicks })
  const range = ctx.spec.axis ?? {}
  const fixed = { ...(range.min !== undefined ? { min: range.min } : {}), ...(range.max !== undefined ? { max: range.max } : {}) }
  const includeZero = o.includeZero ?? true
  if (!ctx.valueAxis && range.min === undefined && range.max === undefined) {
    // No axis to read: let the data fill the plot.
    const lo = includeZero ? Math.min(0, o.lo) : o.lo
    const hi = includeZero ? Math.max(0, o.hi) : o.hi
    return lo === hi ? niceScale(lo, hi, { maxTicks, includeZero }) : { min: lo, max: hi, step: hi - lo, ticks: [] }
  }
  return niceScale(o.lo, o.hi, { maxTicks, includeZero, ...fixed })
}

const zeroOn = (scale: LinearScale): number => Math.min(scale.max, Math.max(scale.min, 0))

/** Values beyond an axis range the author set are drawn at its edge, clipped to the range. */
const clamp = (scale: LinearScale, v: number): number => Math.min(scale.max, Math.max(scale.min, v))

/** Axes with categories along the bottom: column, stacked column, line and waterfall charts. */
export function columnAxes(ctx: ChartContext, area: Box, categories: string[], o: AxisOptions): Axes {
  const font = ctx.font
  const lh = lineHeight(font)
  const n = Math.max(1, categories.length)
  const labelRoom = Math.max(o.roomAbove ? lh + font * 0.3 : 0, ctx.valueAxis ? font * 0.6 : 0)
  const top = labelRoom + (o.extraTop ?? 0)
  const below = o.roomBelow ? lh : 0
  let scale = makeScale(ctx, o, area.h * 0.7, font * 2.6)
  let axisWidth = 0
  let fitted = fitLabels(categories, area.w / n, font)
  let plotH = area.h
  const short = abbreviateMonths(categories)
  for (let pass = 0; pass < 3; pass += 1) {
    const ticks = ctx.valueAxis ? scale.ticks.map((t) => textWidth(tickText(ctx, t, scale.step, Boolean(o.percent)), font)) : []
    axisWidth = ticks.length ? Math.max(...ticks) + font * 0.6 : 0
    const band = (area.w - axisWidth) / n
    fitted = fitLabels(categories, Math.max(font, band - font * 0.5), font, 3)
    // Month names that do not fit are written the way charts usually write them: Jan, Feb, Mar.
    if (short && fitted.wrapped.some((w) => w.truncated)) fitted = fitLabels(short, Math.max(font, band - font * 0.5), font, 3)
    const labelsHeight = fitted.lines * lineHeight(fitted.size) + font * 0.5 + (below ? font * 0.25 : 0)
    plotH = Math.max(font * 3, area.h - top - below - labelsHeight)
    scale = makeScale(ctx, o, plotH, font * 2.6)
  }
  const plot: Box = { x: area.x + axisWidth, y: area.y + top, w: area.w - axisWidth, h: plotH }
  const band = plot.w / n
  const at = (v: number): number => project(scale, clamp(scale, v), plot.y + plot.h, plot.y)
  const base = at(zeroOn(scale))
  const back: string[] = []
  const front: string[] = []
  if (ctx.valueAxis) {
    for (const t of scale.ticks) {
      const y = at(t)
      if (Math.abs(y - base) > 0.5) back.push(line(plot.x, y, plot.x + plot.w, y, INK.grid, 0.5))
      const label = tickText(ctx, t, scale.step, Boolean(o.percent))
      front.push(text(plot.x - font * 0.5, centredBaseline(y, font), label, { size: font, fill: INK.text, anchor: 'end' }))
    }
  }
  front.push(line(plot.x, base, plot.x + plot.w, base, INK.text, 0.75))
  if (fitted.wrapped.some((w) => w.truncated)) warn(ctx, LABELS_CUT)
  const labelTop = plot.y + plot.h + below + font * (below ? 0.7 : 0.45)
  fitted.wrapped.forEach((w, i) => {
    front.push(textLines(plot.x + band * (i + 0.5), labelTop, w.lines, { size: fitted.size, fill: INK.text, anchor: 'middle' }))
  })
  return {
    plot, scale, at, base, band, bandStart: (i) => plot.x + band * i,
    back: group('gridlines', back.join('')), front: group('axes', front.join('')), edge: plot.y - labelRoom,
  }
}

export interface BarAxesOptions extends AxisOptions {
  /** Width needed right of the longest bar for its data label. */
  labelRoom: number
  /** Width needed left of the longest negative bar for its data label. */
  negativeRoom: number
  /** Extra width right of the labels, for annotation arrows. */
  extraRight?: number
}

/** Axes with categories down the left: bar, stacked bar and 100% bar charts. */
export function barAxes(ctx: ChartContext, area: Box, categories: string[], o: BarAxesOptions): Axes {
  const font = ctx.font
  const lh = lineHeight(font)
  const n = Math.max(1, categories.length)
  const axisHeight = ctx.valueAxis ? lh + font * 0.3 : 0
  const plotH = Math.max(font * 2, area.h - axisHeight)
  const band = plotH / n
  const natural = Math.max(...categories.map((c) => textWidth(c, font)), font)
  const labelMax = Math.min(area.w * 0.38, natural)
  const maxLines = Math.min(3, Math.max(1, Math.floor((band * 0.95) / lh)))
  const fitted = fitLabels(categories, labelMax, font, maxLines)
  if (fitted.wrapped.some((w) => w.truncated)) warn(ctx, LABELS_CUT)
  const labelWidth = Math.max(...fitted.wrapped.map((w) => w.width), 0) + font * 0.6
  let scale = makeScale(ctx, o, area.w * 0.5, font * 5)
  let plot: Box = { x: area.x, y: area.y, w: area.w, h: plotH }
  for (let pass = 0; pass < 2; pass += 1) {
    const last = scale.ticks[scale.ticks.length - 1]
    const tickHalf = ctx.valueAxis && last !== undefined ? textWidth(tickText(ctx, last, scale.step, Boolean(o.percent)), font) / 2 : 0
    const left = labelWidth + o.negativeRoom
    const right = Math.max(o.labelRoom + (o.extraRight ?? 0), tickHalf)
    plot = { x: area.x + left, y: area.y, w: Math.max(font * 4, area.w - left - right), h: plotH }
    const widest = Math.max(...scale.ticks.map((t) => textWidth(tickText(ctx, t, scale.step, Boolean(o.percent)), font)), font)
    scale = makeScale(ctx, o, plot.w, widest + font * 2)
  }
  const at = (v: number): number => project(scale, clamp(scale, v), plot.x, plot.x + plot.w)
  const base = at(zeroOn(scale))
  const front: string[] = []
  if (ctx.valueAxis) {
    for (const t of scale.ticks) {
      const label = tickText(ctx, t, scale.step, Boolean(o.percent))
      front.push(text(at(t), plot.y + plot.h + font * 0.3 + font * 0.86, label, { size: font, fill: INK.text, anchor: 'middle' }))
    }
  }
  front.push(line(base, plot.y, base, plot.y + plot.h, INK.text, 0.75))
  fitted.wrapped.forEach((w, i) => {
    const cy = plot.y + band * (i + 0.5)
    front.push(textLinesCentred(area.x + labelWidth - font * 0.6, cy, w.lines, { size: fitted.size, fill: INK.text, anchor: 'end' }))
  })
  return {
    plot, scale, at, base, band, bandStart: (i) => plot.y + band * i, back: '', front: group('axes', front.join('')),
    edge: plot.x + plot.w + o.labelRoom,
  }
}
