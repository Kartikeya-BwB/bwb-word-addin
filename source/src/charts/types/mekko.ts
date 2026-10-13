// Mekko (marimekko) charts: one column per category, as wide as the category's share of the grand total, with its
// segments stacked to 100%. Segments carry their share, columns their total above, categories their name below.

import { autoDecimals, formatNumber } from '../format.ts'
import type { Box, ChartContext, LegendItem } from '../frame.ts'
import { LABELS_CUT, chartFrame, fmt, warn } from '../frame.ts'
import { CAP_HEIGHT, MIN_FONT_PT, fitLabels, lineHeight, textWidth } from '../measure.ts'
import { INK, inkOn } from '../palette.ts'
import type { DrawSeries, Prepared } from '../series.ts'
import { prepareSeries } from '../series.ts'
import { centredBaseline, group, line, rect, text, textLines } from '../svg.ts'
import { placeholder } from './empty.ts'

export interface MekkoSegment {
  series: DrawSeries
  value: number
  /** Share of the column, 0 to 1. */
  share: number
  /** Bottom and top of the segment as shares of the column height. */
  from: number
  to: number
}

export interface MekkoColumn {
  category: string
  total: number
  /** Left edge and width as shares of the plot width, before gaps. */
  start: number
  width: number
  segments: MekkoSegment[]
}

/** Columns with widths in proportion to category totals; negative values count as zero. */
export function mekkoColumns(p: Prepared): MekkoColumn[] {
  const totals = p.categories.map((_, c) => p.series.reduce((sum, s) => sum + Math.max(0, s.values[c] ?? 0), 0))
  const grand = totals.reduce((a, b) => a + b, 0)
  let start = 0
  return p.categories.map((category, c) => {
    const total = totals[c] ?? 0
    const width = grand > 0 ? total / grand : 0
    let from = 0
    const segments: MekkoSegment[] = []
    for (const s of p.series) {
      const value = Math.max(0, s.values[c] ?? 0)
      if (value <= 0 || total <= 0) continue
      const share = value / total
      segments.push({ series: s, value, share, from, to: from + share })
      from += share
    }
    const column = { category, total, start, width, segments }
    start += width
    return column
  })
}

export function renderMekko(ctx: ChartContext): string {
  const p = prepareSeries(ctx, { sortable: true, highlight: false })
  if (p.series.some((s) => s.values.some((v) => v !== null && v < 0))) warn(ctx, 'Mekko charts leave out negative values.')
  const columns = mekkoColumns(p).filter((c) => c.width > 0)
  if (!columns.length) return placeholder(ctx, 'Add positive numbers to the data to draw the Mekko chart.')
  const font = ctx.font
  const items: LegendItem[] = p.series.map((s) => ({ label: s.name, colour: s.colour }))
  const frame = chartFrame(ctx, items)
  const lh = lineHeight(font)
  const gapPt = Math.max(1, font * 0.15)
  const axisWidth = ctx.valueAxis ? textWidth('100%', font) + font * 0.6 : 0
  const plotW = frame.plot.w - axisWidth
  const available = plotW - gapPt * (columns.length - 1)
  const widths = columns.map((c) => available * c.width)
  const fitted = fitLabels(columns.map((c) => c.category), widths.map((w) => Math.max(1, w - font * 0.3)), font, 3)
  if (fitted.wrapped.some((w) => w.truncated)) warn(ctx, LABELS_CUT)
  const labelsHeight = fitted.lines * lineHeight(fitted.size) + font * 0.5
  const top = frame.plot.y + (ctx.labels ? lh : 0)
  const bottom = frame.plot.y + frame.plot.h - labelsHeight
  const plot: Box = { x: frame.plot.x + axisWidth, y: top, w: plotW, h: Math.max(font * 3, bottom - top) }
  const yAt = (share: number): number => plot.y + plot.h - share * plot.h
  const shapes: string[] = []
  const labels: string[] = []
  const decimals = ctx.format.decimals ?? autoDecimals(columns.map((c) => c.total))
  const totalTexts = columns.map((c) => fmt(ctx, c.total, { decimals }))
  let totalSize: number | null = null
  for (let s = font; s >= MIN_FONT_PT && totalSize === null; s -= 0.5) {
    if (totalTexts.every((t, i) => textWidth(t, s, true) <= (widths[i] ?? 0) + gapPt)) totalSize = s
  }
  let x = plot.x
  columns.forEach((column, i) => {
    const w = widths[i] ?? 0
    column.segments.forEach((g, k) => {
      const yTop = yAt(g.to)
      const yBottom = yAt(g.from)
      shapes.push(rect(x, yTop, w, yBottom - yTop, g.series.colour))
      if (k > 0) shapes.push(line(x, yBottom, x + w, yBottom, INK.white, 0.75))
      const label = formatNumber(g.share * 100, { percent: true }, { decimals: 0 })
      if (ctx.labels && textWidth(label, font) <= w - font * 0.3 && yBottom - yTop >= font * (CAP_HEIGHT + 0.5)) {
        labels.push(text(x + w / 2, centredBaseline((yTop + yBottom) / 2, font), label, { size: font, fill: inkOn(g.series.colour), anchor: 'middle' }))
      }
    })
    if (ctx.labels && totalSize !== null) {
      labels.push(text(x + w / 2, plot.y - totalSize * 0.35, totalTexts[i] ?? '', { size: totalSize, fill: INK.data, anchor: 'middle', bold: true }))
    }
    const wrapped = fitted.wrapped[i]
    if (wrapped) labels.push(textLines(x + w / 2, plot.y + plot.h + font * 0.45, wrapped.lines, { size: fitted.size, fill: INK.text, anchor: 'middle' }))
    x += w + gapPt
  })
  if (ctx.labels && totalSize === null) warn(ctx, 'Some column totals did not fit above their columns and are left out.')
  const axis: string[] = []
  if (ctx.valueAxis) {
    for (const t of [0, 25, 50, 75, 100]) {
      const y = yAt(t / 100)
      axis.push(text(plot.x - font * 0.5, centredBaseline(y, font), `${t}%`, { size: font, fill: INK.text, anchor: 'end' }))
    }
  }
  axis.push(line(plot.x, plot.y + plot.h, plot.x + plot.w, plot.y + plot.h, INK.text, 0.75))
  return frame.svg + group('segments', shapes.join('')) + group('axes', axis.join('')) + group('labels', labels.join(''))
}
