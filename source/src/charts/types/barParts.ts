// Pieces shared by column and bar charts: stacking, label sizes, and the widest a bar gets.

import { autoDecimals, formatNumber } from '../format.ts'
import type { ChartContext, LegendItem } from '../frame.ts'
import { fmt, unitsNote, warn } from '../frame.ts'
import { MIN_FONT_PT, lineHeight, textWidth } from '../measure.ts'
import type { DrawSeries, Prepared } from '../series.ts'

export type BarMode = 'clustered' | 'stacked' | 'percent'

export interface Segment {
  series: DrawSeries
  /** Value drawn: the number, or its share for 100% charts. */
  value: number
  from: number
  to: number
  colour: string
}

export interface Stack {
  segments: Segment[]
  total: number
  top: number
  bottom: number
}

export function stacks(ctx: ChartContext, p: Prepared, mode: BarMode): Stack[] {
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

export const legendItems = (p: Prepared): LegendItem[] => p.series.map((s) => ({ label: s.name, colour: s.colour }))

/** The largest label size from the chart's size down to 8 pt at which every text fits its width, or null. */
export function fitSize(texts: string[], room: number, start: number): number | null {
  for (let size = start; size >= MIN_FONT_PT; size -= 0.5) {
    if (texts.every((t) => textWidth(t, size) <= room)) return size
  }
  return null
}

export interface LabelPlan {
  /** Labels without currency and scale, with the units written once at the top left. */
  bare: boolean
  decimals: number
}

/** Leaves out the units from labels when the value axis shows numbers, or when labels with units would not fit. */
export function planLabels(ctx: ChartContext, values: number[], room: number): LabelPlan {
  const decimals = ctx.format.decimals ?? autoDecimals(values)
  if (!unitsNote(ctx.format)) return { bare: false, decimals }
  if (ctx.valueAxis) return { bare: true, decimals }
  const full = values.map((v) => fmt(ctx, v, { decimals }))
  return { bare: fitSize(full, room, ctx.font) === null, decimals }
}

export const shareText = (value: number): string => formatNumber(value, { percent: true }, { decimals: 0 })

/** The widest a bar (or a cluster of bars) gets, so a chart of a few categories does not turn into blocks. */
export function maxBar(ctx: ChartContext, clusters: number): number {
  return ctx.font * (clusters > 1 ? 3.5 * clusters + 1 : 7)
}

/** Label size for labels stacked a pitch apart, or null when even 8 pt labels would touch. */
export function sizeForPitch(pitch: number, font: number): number | null {
  for (let size = font; size >= MIN_FONT_PT; size -= 0.5) if (lineHeight(size) * 0.85 <= pitch) return size
  return null
}
