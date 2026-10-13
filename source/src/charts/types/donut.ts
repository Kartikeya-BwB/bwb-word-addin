// Donut charts: parts of one whole, at most six segments (the smallest are combined as "Other" beyond that), clockwise
// from twelve o'clock, with shares on the ring, the total in the middle and a legend of names and values beside it.

import { COLOURS } from '../../brand/tokens.ts'
import { autoDecimals, formatNumber } from '../format.ts'
import type { ChartContext } from '../frame.ts'
import { chartFrame, fmt, warn } from '../frame.ts'
import { CAP_HEIGHT, MIN_FONT_PT, ellipsise, lineHeight, textWidth, wrapText } from '../measure.ts'
import { INK, inkOn, paletteColour } from '../palette.ts'
import { prepareSeries } from '../series.ts'
import { centredBaseline, group, num, path, rect, text, textLinesCentred } from '../svg.ts'
import { placeholder } from './empty.ts'

export const MAX_SEGMENTS = 6

export interface DonutSegment {
  label: string
  value: number
  share: number
  colour: string
}

/** Segments in data order, the smallest folded into "Other" when there are more than six. */
export function donutSegments(labels: string[], values: Array<number | null>, colourOf: (i: number) => string): DonutSegment[] {
  let items = labels.map((label, i) => ({ label, value: values[i] ?? 0, i })).filter((d) => d.value > 0)
  if (items.length > MAX_SEGMENTS) {
    const kept = new Set([...items].sort((a, b) => b.value - a.value).slice(0, MAX_SEGMENTS - 1).map((d) => d.i))
    const other = items.filter((d) => !kept.has(d.i)).reduce((sum, d) => sum + d.value, 0)
    items = [...items.filter((d) => kept.has(d.i)), { label: 'Other', value: other, i: MAX_SEGMENTS - 1 }]
  }
  const total = items.reduce((sum, d) => sum + d.value, 0)
  return items.map((d, k) => ({ label: d.label, value: d.value, share: total > 0 ? d.value / total : 0, colour: colourOf(k) }))
}

const point = (cx: number, cy: number, r: number, turn: number): [number, number] => [
  cx + r * Math.sin(turn * 2 * Math.PI),
  cy - r * Math.cos(turn * 2 * Math.PI),
]

/** An annular sector from one turn fraction to another, clockwise from twelve o'clock. */
export function sector(cx: number, cy: number, outer: number, inner: number, from: number, to: number): string {
  if (to - from >= 0.9999) {
    // A whole ring: two half arcs each way, with the hole cut out by the even-odd rule.
    return (
      `M${num(cx)} ${num(cy - outer)}A${num(outer)} ${num(outer)} 0 1 1 ${num(cx)} ${num(cy + outer)}` +
      `A${num(outer)} ${num(outer)} 0 1 1 ${num(cx)} ${num(cy - outer)}Z` +
      `M${num(cx)} ${num(cy - inner)}A${num(inner)} ${num(inner)} 0 1 0 ${num(cx)} ${num(cy + inner)}` +
      `A${num(inner)} ${num(inner)} 0 1 0 ${num(cx)} ${num(cy - inner)}Z`
    )
  }
  const large = to - from > 0.5 ? 1 : 0
  const [x1, y1] = point(cx, cy, outer, from)
  const [x2, y2] = point(cx, cy, outer, to)
  const [x3, y3] = point(cx, cy, inner, to)
  const [x4, y4] = point(cx, cy, inner, from)
  return (
    `M${num(x1)} ${num(y1)}A${num(outer)} ${num(outer)} 0 ${large} 1 ${num(x2)} ${num(y2)}` +
    `L${num(x3)} ${num(y3)}A${num(inner)} ${num(inner)} 0 ${large} 0 ${num(x4)} ${num(y4)}Z`
  )
}

export function renderDonut(ctx: ChartContext): string {
  const p = prepareSeries(ctx, { sortable: true, highlight: true })
  const first = p.series[0]
  if (!first) return placeholder(ctx, 'Add a column of values to draw the donut.')
  if (p.series.length > 1) warn(ctx, 'A donut shows the first series only.')
  if (first.values.some((v) => v !== null && v < 0)) warn(ctx, 'Donuts leave out negative values.')
  // Segments take the palette in order; with a highlight, one segment is green and the rest navy.
  const colourOf = (k: number): string => (p.highlighted >= 0 ? COLOURS.navy : paletteColour(k))
  const segments = donutSegments(p.categories, first.values, colourOf)
  if (p.highlighted >= 0) {
    const name = p.categories[p.highlighted]
    for (const s of segments) if (s.label === name) s.colour = COLOURS.green
  }
  if (!segments.length) return placeholder(ctx, 'Add positive values to draw the donut.')
  if (p.categories.length > MAX_SEGMENTS) warn(ctx, `Donuts show ${MAX_SEGMENTS} segments at most; the smallest are combined as \u{201C}Other\u{201D}.`)
  const font = ctx.font
  const lh = lineHeight(font)
  const frame = chartFrame(ctx, [])
  const plot = frame.plot
  const decimals = ctx.format.decimals ?? autoDecimals(segments.map((s) => s.value))
  const values = segments.map((s) => fmt(ctx, s.value, { decimals }))
  const showValues = !ctx.format.percent
  // The legend lists swatch, name, and the value when shares are not already the values. It sits beside the ring,
  // or below it when that leaves a bigger ring, as in a narrow chart.
  const swatch = font * 0.75
  const valueWidth = showValues ? Math.max(...values.map((v) => textWidth(v, font, true))) : 0
  const nameMax = Math.max(font * 4, Math.min(plot.w * 0.4, Math.max(...segments.map((s) => textWidth(s.label, font)))))
  const legendWidth = ctx.legend ? swatch + font * 0.5 + nameMax + (showValues ? font * 1.2 + valueWidth : 0) : 0
  const pitch = lh * 1.25
  const legendHeight = ctx.legend ? segments.length * pitch : 0
  const gap = ctx.legend ? font * 2 : 0
  const beside = Math.min(plot.h, plot.w - legendWidth - gap)
  const below = Math.min(plot.w, plot.h - legendHeight - font)
  const under = ctx.legend && below > beside * 1.15
  const diameter = Math.max(font * 4, (under ? below : beside) - 2)
  const outer = diameter / 2
  const inner = outer * 0.6
  let cx = plot.x + plot.w / 2
  let cy = plot.y + plot.h / 2
  let legendX = cx - legendWidth / 2
  let legendTop = cy
  if (under) {
    const top = plot.y + Math.max(0, (plot.h - diameter - font - legendHeight) / 2)
    cy = top + outer
    legendTop = top + diameter + font
  } else {
    cx = plot.x + Math.max(0, (plot.w - diameter - gap - legendWidth) / 2) + outer
    legendX = cx + outer + gap
    legendTop = cy - legendHeight / 2
  }
  const shapes: string[] = []
  const labels: string[] = []
  let turn = 0
  segments.forEach((s, k) => {
    const from = turn
    const to = turn + s.share
    turn = to
    shapes.push(path(sector(cx, cy, outer, inner, from, to), s.colour, segments.length > 1 ? { colour: INK.white, width: Math.max(1, font * 0.12) } : undefined))
    if (!ctx.labels) return
    const label = formatNumber(s.share * 100, { percent: true }, { decimals: 0 })
    const mid = (outer + inner) / 2
    const arc = s.share * 2 * Math.PI * mid
    if (textWidth(label, font) + font * 0.6 <= Math.min(arc, (outer - inner) * 1.8) && outer - inner >= font * (CAP_HEIGHT + 0.6)) {
      const [lx, ly] = point(cx, cy, mid, (from + to) / 2)
      labels.push(text(lx, centredBaseline(ly, font), label, { size: font, fill: inkOn(s.colour), anchor: 'middle' }))
    }
    if (ctx.legend) {
      const rowTop = legendTop + k * pitch
      const lx = legendX
      shapes.push(rect(lx, rowTop + (lh - swatch) / 2, swatch, swatch, s.colour))
      const name = ellipsise(s.label, nameMax, font)
      labels.push(text(lx + swatch + font * 0.5, centredBaseline(rowTop + lh / 2, font), name, { size: font, fill: INK.text }))
      if (showValues) {
        labels.push(text(lx + legendWidth, centredBaseline(rowTop + lh / 2, font), values[k] ?? '', { size: font, fill: INK.data, anchor: 'end', bold: true }))
      }
    }
  })
  const total = segments.reduce((sum, s) => sum + s.value, 0)
  const centre = ctx.spec.centreLabel ?? (ctx.format.percent ? '' : fmt(ctx, total, { decimals }))
  if (centre.trim()) {
    const room = inner * 1.5
    let size = Math.min(font * 2, Math.max(MIN_FONT_PT, font * 1.6))
    while (size > MIN_FONT_PT && wrapText(centre, room, size, 2, true).truncated) size -= 0.5
    const lines = wrapText(centre, room, size, 2, true).lines
    labels.push(textLinesCentred(cx, cy, lines, { size, fill: INK.data, anchor: 'middle', bold: true }))
  }
  return frame.svg + group('segments', shapes.join('')) + group('labels', labels.join(''))
}
