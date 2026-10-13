// What every chart shares: the resolved options, the legend at the top left (not boxed), an optional units note,
// and the optional source line at the bottom left in 8 pt charcoal. The chart body gets the box that is left.

import type { NumberFormat } from './format.ts'
import { formatNumber, unitsLabel } from './format.ts'
import { MIN_FONT_PT, ellipsise, lineHeight, textWidth, wrapText } from './measure.ts'
import { INK } from './palette.ts'
import type { ChartResources, ChartSpec } from './spec.ts'
import { centredBaseline, circle, diamond, group, line, rect, text, textLines } from './svg.ts'

export interface Box {
  x: number
  y: number
  w: number
  h: number
}

export interface ChartContext {
  spec: ChartSpec
  width: number
  height: number
  /** Size of labels, legend and axis text, in points. */
  font: number
  labels: boolean
  legend: boolean
  valueAxis: boolean
  /** Gap between bars as a share of the bar width. */
  gap: number
  format: NumberFormat
  warnings: string[]
  /** Set when data labels were left out for lack of room, so the chart can show a value axis instead. */
  labelsDropped: boolean
  /** Data loaded by the caller, such as the world map. */
  resources: ChartResources
}

export const LABELS_DROPPED = 'Some data labels did not fit and are left out. Make the chart bigger or turn labels off.'
export const LABELS_TO_AXIS = 'Data labels did not fit, so the chart shows a value axis instead. Make the chart bigger to label the values.'
export const LABELS_CUT = 'Some labels are cut short to fit. Make the chart bigger or the labels shorter.'

/** Records that data labels were left out. */
export function dropLabels(ctx: ChartContext): void {
  ctx.labelsDropped = true
  warn(ctx, LABELS_DROPPED)
}

export const SOURCE_PT = 8

/** Label size for a chart width: 12 pt across a slide's content area, 9 pt across an A4 page, never below 8 pt. */
export function defaultFontPt(widthPt: number): number {
  return widthPt >= 640 ? 12 : widthPt >= 400 ? 9 : MIN_FONT_PT
}

export function warn(ctx: ChartContext, message: string): void {
  if (!ctx.warnings.includes(message)) ctx.warnings.push(message)
}

/** A value written in the chart's format; `bare` leaves out currency and scale when a units note carries them. */
export function fmt(ctx: ChartContext, value: number, options: { decimals?: number; plus?: boolean; bare?: boolean } = {}): string {
  return formatNumber(value, ctx.format, options)
}

export type Mark = 'square' | 'line' | 'dashed' | 'marker' | 'diamond'

export interface LegendItem {
  label: string
  colour: string
  mark?: Mark
}

function legendMark(item: LegendItem, x: number, cy: number, size: number): { svg: string; width: number } {
  const mark = item.mark ?? 'square'
  if (mark === 'square') {
    const side = size * 0.75
    return { svg: rect(x, cy - side / 2, side, side, item.colour), width: side }
  }
  if (mark === 'diamond') return { svg: diamond(x + size * 0.45, cy, size * 0.45, item.colour), width: size * 0.9 }
  const long = size * 1.6
  const stroke = Math.max(1.25, size * 0.17)
  const svg = line(x, cy, x + long, cy, item.colour, stroke, mark === 'dashed' ? `${num2(stroke * 2.5)} ${num2(stroke * 1.6)}` : undefined)
  const dot = mark === 'marker' ? circle(x + long / 2, cy, size * 0.26, item.colour, { colour: INK.white, width: size * 0.08 }) : ''
  return { svg: svg + dot, width: long }
}

const num2 = (v: number): string => String(Math.round(v * 100) / 100)

/** Legend items in rows from the top left, wrapping when a row is full. */
export function legendBlock(items: LegendItem[], x: number, y: number, maxWidth: number, size: number, note = ''): { svg: string; height: number } {
  const parts: string[] = []
  const lh = lineHeight(size)
  const gapAfterMark = size * 0.45
  const gapBetween = size * 1.4
  let cx = x
  let row = 0
  const place = (width: number): number => {
    if (cx > x && cx + width > x + maxWidth) {
      row += 1
      cx = x
    }
    const at = cx
    cx += width + gapBetween
    return at
  }
  const centreY = (): number => y + row * lh + lh / 2
  if (note) {
    const label = ellipsise(note, maxWidth, size)
    const at = place(textWidth(label, size))
    parts.push(text(at, centredBaseline(centreY(), size), label, { size, fill: INK.text }))
  }
  for (const item of items) {
    const sample = legendMark(item, 0, 0, size)
    const label = ellipsise(item.label, Math.max(size * 3, maxWidth - sample.width - gapAfterMark), size)
    const at = place(sample.width + gapAfterMark + textWidth(label, size))
    const cy = centreY()
    parts.push(legendMark(item, at, cy - size * 0.04, size).svg)
    parts.push(text(at + sample.width + gapAfterMark, centredBaseline(cy, size), label, { size, fill: INK.text }))
  }
  return { svg: group('legend', parts.join('')), height: parts.length ? (row + 1) * lh : 0 }
}

/** "Source: ..." in 8 pt italic charcoal at the bottom left, at most two lines. */
export function sourceBlock(source: string, x: number, bottom: number, maxWidth: number): { svg: string; height: number } {
  const trimmed = source.trim()
  if (!trimmed) return { svg: '', height: 0 }
  const label = /^sources?\s*:/i.test(trimmed) ? trimmed : `Source: ${trimmed}`
  const wrapped = wrapText(label, maxWidth, SOURCE_PT, 2)
  const height = wrapped.lines.length * lineHeight(SOURCE_PT)
  const svg = textLines(x, bottom - height, wrapped.lines, { size: SOURCE_PT, fill: INK.text, italic: true })
  return { svg: group('source', svg), height }
}

export interface Frame {
  plot: Box
  svg: string
}

/**
 * Lays out the legend (when on and there is more than one item to tell apart), a units note, and the source line,
 * and returns the box left for the chart itself.
 */
export function chartFrame(ctx: ChartContext, items: LegendItem[], note = ''): Frame {
  const showLegend = ctx.legend && items.length > 1
  const legend = legendBlock(showLegend ? items : [], 0, 0, ctx.width, ctx.font, note)
  const source = ctx.spec.sourceInChart && ctx.spec.source ? sourceBlock(ctx.spec.source, 0, ctx.height, ctx.width) : { svg: '', height: 0 }
  const top = legend.height ? legend.height + ctx.font * 0.8 : 0
  const bottom = source.height ? source.height + ctx.font * 0.7 : 0
  return {
    plot: { x: 0, y: top, w: ctx.width, h: Math.max(ctx.font * 4, ctx.height - top - bottom) },
    svg: legend.svg + source.svg,
  }
}

/** The units a chart's numbers are in, written once ("USD m") when labels and ticks leave them out. */
export function unitsNote(format: NumberFormat): string {
  return format.percent ? '' : unitsLabel(format)
}
