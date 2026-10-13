// Annotations for column, bar and line charts: a growth arrow (compound annual growth rate) and a difference arrow
// between two categories, an average line, a target line, and a break in one bar far taller than the rest.
// Arrows are thin navy brackets in lanes beyond the data, with the label in a small gap in the line; the average is a
// dashed charcoal line and the target a green one, each with a small charcoal label where it is free.

import { parseDate } from './dates.ts'
import { formatNumber } from './format.ts'
import type { Box, ChartContext } from './frame.ts'
import { fmt, warn } from './frame.ts'
import { lineHeight, textWidth } from './measure.ts'
import { INK } from './palette.ts'
import type { Rect } from './place.ts'
import { overlap, textRect } from './place.ts'
import { centredBaseline, group, line, num, path, rect, text } from './svg.ts'

export interface Annotations {
  /** Compound annual growth between two categories, as an arrow above them. */
  cagr?: { from: string; to: string }
  /** The difference between two categories, in the chart's units or in per cent. */
  difference?: { from: string; to: string; percent?: boolean }
  /** A dashed line at the average of the categories. */
  average?: boolean
  /** A line at a target value, labelled. */
  target?: { value: number; label?: string }
  /** Shortens the one bar far taller than the rest, with a break mark; its label keeps the true value. */
  breakOutlier?: boolean
}

/** Growth per period that turns start into end over `periods`; null when it cannot be worked out. */
export function cagr(start: number, end: number, periods: number): number | null {
  if (!(start > 0) || end < 0 || !(periods > 0) || !Number.isFinite(start + end + periods)) return null
  return (end / start) ** (1 / periods) - 1
}

const YEAR = /^(?:FY|CY)?\s?'?(\d{4}|\d{2})[A-Z]?$/i

/** Years between two categories: from their years or dates when they are periods, otherwise one per step. */
export function periodsBetween(fromLabel: string, toLabel: string, fromIndex: number, toIndex: number): number {
  const year = (label: string): number | null => {
    const m = YEAR.exec(label.trim())
    if (!m?.[1]) return null
    const y = Number(m[1])
    return m[1].length === 2 ? 2000 + y : y
  }
  const [a, b] = [year(fromLabel), year(toLabel)]
  if (a !== null && b !== null) return b - a
  const [da, db] = [parseDate(fromLabel), parseDate(toLabel)]
  if (da && db) return Math.round(((db.first - da.first) / 365.2425) * 100) / 100
  return toIndex - fromIndex
}

/** "CAGR +8.4%", with a true minus sign for decline. */
export const cagrLabel = (rate: number): string => `CAGR ${formatNumber(rate * 100, { percent: true }, { decimals: 1, plus: true })}`

export interface Bracket {
  from: number
  to: number
  label: string
}

export interface AnnotationPlan {
  brackets: Bracket[]
  average: number | null
  target: { value: number; label: string } | null
  /** Room the brackets need beyond the data: above columns and lines, right of bars. */
  room: number
  /** Size of annotation labels. */
  size: number
}

const same = (a: string, b: string): boolean => a.trim().toLowerCase() === b.trim().toLowerCase()

/**
 * Works out what to draw from the chart's categories and the level of each (a stack's total, or the first series),
 * and how much room the arrows need.
 */
export function planAnnotations(ctx: ChartContext, categories: string[], levels: Array<number | null>, horizontal: boolean): AnnotationPlan {
  const a = ctx.spec.annotations ?? {}
  const size = Math.max(8, ctx.font * 0.9)
  const brackets: Bracket[] = []
  const find = (name: string): number => categories.findIndex((c) => same(c, name))
  const pair = (kind: string, from: string, to: string): [number, number] | null => {
    const [i, j] = [find(from), find(to)]
    if (i < 0 || j < 0 || i === j) {
      warn(ctx, `The ${kind} needs two different categories from the data.`)
      return null
    }
    const [vi, vj] = [levels[i] ?? null, levels[j] ?? null]
    if (vi === null || vj === null) {
      warn(ctx, `The ${kind} needs values at both ends.`)
      return null
    }
    return [i, j]
  }
  if (a.cagr) {
    const ends = pair('growth arrow', a.cagr.from, a.cagr.to)
    if (ends) {
      const [i, j] = ends
      const [lo, hi] = i < j ? [i, j] : [j, i]
      const rate = cagr(levels[lo] ?? 0, levels[hi] ?? 0, periodsBetween(categories[lo] ?? '', categories[hi] ?? '', lo, hi))
      if (rate === null) warn(ctx, 'The growth rate needs a value above zero at the start and periods in order.')
      else brackets.push({ from: lo, to: hi, label: cagrLabel(rate) })
    }
  }
  if (a.difference) {
    const ends = pair('difference arrow', a.difference.from, a.difference.to)
    if (ends) {
      const [i, j] = ends
      const [start, end] = [levels[i] ?? 0, levels[j] ?? 0]
      if (a.difference.percent && start === 0) warn(ctx, 'A difference in per cent needs a starting value other than zero.')
      else {
        const change = a.difference.percent ? ((end - start) / Math.abs(start)) * 100 : end - start
        const label = a.difference.percent
          ? formatNumber(change, { percent: true }, { decimals: Math.abs(change) >= 10 ? 0 : 1, plus: true })
          : fmt(ctx, change, { plus: true })
        brackets.unshift({ from: i, to: j, label })
      }
    }
  }
  const known = levels.filter((v): v is number => v !== null)
  const average = a.average && known.length ? known.reduce((s, v) => s + v, 0) / known.length : null
  const target = a.target && Number.isFinite(a.target.value)
    ? { value: a.target.value, label: `${a.target.label?.trim() || 'Target'}: ${fmt(ctx, a.target.value)}` }
    : null
  const lane = horizontal
    ? (b: Bracket): number => textWidth(b.label, size, true) + ctx.font * 1.4
    : (): number => lineHeight(size) + ctx.font * 0.5
  return { brackets, average, target, room: brackets.reduce((sum, b) => sum + lane(b), 0), size }
}

export interface AnnotationGeometry {
  horizontal: boolean
  /** Centre of a category along the category axis. */
  centre: (index: number) => number
  /** Outer edge of each category's data and label along the value axis: the top for columns, the right for bars. */
  reach: Array<number | null>
  /** A value's position along the value axis. */
  at: (value: number) => number
  plot: Box
  /** Where the arrow lanes start: just beyond the data labels. */
  laneStart: number
}

function arrowhead(x: number, y: number, size: number, direction: 'down' | 'left'): string {
  const a = Math.max(1.8, size * 0.24)
  const d = direction === 'down'
    ? `M${num(x - a)} ${num(y - a * 1.6)}L${num(x + a)} ${num(y - a * 1.6)}L${num(x)} ${num(y)}Z`
    : `M${num(x + a * 1.6)} ${num(y - a)}L${num(x + a * 1.6)} ${num(y + a)}L${num(x)} ${num(y)}Z`
  return path(d, INK.navy)
}

/** A label beside a reference line: tried at the end, above then below, then at the start; left out if all are taken. */
function lineLabel(ctx: ChartContext, g: AnnotationGeometry, position: number, label: string, size: number, taken: Rect[]): string {
  const gap = size * 0.35
  const { x, y, w, h } = g.plot
  const spots: Array<{ x: number; y: number; anchor: 'start' | 'end' }> = g.horizontal
    ? [
        { x: position + gap, y: y + size * 0.86, anchor: 'start' },
        { x: position - gap, y: y + size * 0.86, anchor: 'end' },
        { x: position + gap, y: y + h - size * 0.3, anchor: 'start' },
        { x: position - gap, y: y + h - size * 0.3, anchor: 'end' },
      ]
    : [
        { x: x + w, y: position - gap, anchor: 'end' },
        { x: x + w, y: position + gap + size * 0.72, anchor: 'end' },
        { x, y: position - gap, anchor: 'start' },
        { x, y: position + gap + size * 0.72, anchor: 'start' },
      ]
  for (const s of spots) {
    const box = textRect(s.x, s.y, label, size, s.anchor)
    if (box.x0 < x - 0.5 || box.x1 > x + w + 0.5 || box.y0 < y - size * 1.5 || taken.some((t) => overlap(box, t, 0.5))) continue
    taken.push(box)
    return text(s.x, s.y, label, { size, fill: INK.text, anchor: s.anchor })
  }
  warn(ctx, `There was no free space for the label \u{201C}${label}\u{201D}; the line is drawn without it.`)
  return ''
}

/** Draws the planned annotations. `taken` lists the data labels and bars, so line labels avoid them. */
export function drawAnnotations(ctx: ChartContext, plan: AnnotationPlan, g: AnnotationGeometry, taken: Rect[]): string {
  const parts: string[] = []
  const size = plan.size
  const lh = lineHeight(size)
  const gap = ctx.font * 0.25
  for (const [value, label, colour, dash] of [
    [plan.average, plan.average === null ? '' : `Average: ${fmt(ctx, plan.average)}`, INK.charcoal, '3 2'],
    [plan.target?.value ?? null, plan.target?.label ?? '', INK.green, undefined],
  ] as const) {
    if (value === null) continue
    const at = g.at(value)
    parts.push(g.horizontal ? line(at, g.plot.y, at, g.plot.y + g.plot.h, colour, dash ? 0.75 : 1, dash) : line(g.plot.x, at, g.plot.x + g.plot.w, at, colour, dash ? 0.75 : 1, dash))
    parts.push(lineLabel(ctx, g, at, label, size, taken))
  }
  let offset = 0
  plan.brackets.forEach((b) => {
    const target = b.to
    const ends = [b.from, b.to]
    if (!g.horizontal) {
      const laneY = g.laneStart - offset - (lh + ctx.font * 0.5) / 2
      offset += lh + ctx.font * 0.5
      const [x1, x2] = ends.map((i) => g.centre(i))
      for (const i of ends) {
        const cx = g.centre(i)
        const end = (g.reach[i] ?? g.laneStart) - gap
        parts.push(line(cx, laneY, cx, end, INK.navy, 0.75))
        if (i === target) parts.push(arrowhead(cx, end, ctx.font, 'down'))
      }
      parts.push(line(Math.min(x1 ?? 0, x2 ?? 0), laneY, Math.max(x1 ?? 0, x2 ?? 0), laneY, INK.navy, 0.75))
      const mid = ((x1 ?? 0) + (x2 ?? 0)) / 2
      const w = textWidth(b.label, size, true) + size * 0.8
      parts.push(rect(mid - w / 2, laneY - lh / 2, w, lh, INK.white))
      parts.push(text(mid, centredBaseline(laneY, size), b.label, { size, fill: INK.navy, anchor: 'middle', bold: true }))
      return
    }
    const laneX = g.laneStart + offset + ctx.font * 0.7
    offset += textWidth(b.label, size, true) + ctx.font * 1.4
    const [y1, y2] = ends.map((i) => g.centre(i))
    for (const i of ends) {
      const cy = g.centre(i)
      const end = (g.reach[i] ?? g.laneStart) + gap
      parts.push(line(laneX, cy, end, cy, INK.navy, 0.75))
      if (i === target) parts.push(arrowhead(end, cy, ctx.font, 'left'))
    }
    parts.push(line(laneX, Math.min(y1 ?? 0, y2 ?? 0), laneX, Math.max(y1 ?? 0, y2 ?? 0), INK.navy, 0.75))
    parts.push(text(laneX + ctx.font * 0.4, centredBaseline(((y1 ?? 0) + (y2 ?? 0)) / 2, size), b.label, { size, fill: INK.navy, bold: true }))
  })
  return group('annotations', parts.join(''))
}

/** The bar to shorten: the tallest, when it is at least twice the next. The axis then stops a little above the next. */
export function outlierCap(values: Array<number | null>): { index: number; cap: number } | null {
  const ranked = values.map((v, i) => ({ v: v ?? 0, i })).filter((x) => x.v > 0).sort((a, b) => b.v - a.v)
  const [first, second] = ranked
  if (!first || !second || first.v < second.v * 2) return null
  return { index: first.i, cap: second.v * 1.35 }
}

/** A white slanted cut across a bar near its far end, showing it is not drawn to scale. */
export function breakMark(horizontal: boolean, x: number, y: number, w: number, h: number, font: number): string {
  const t = Math.max(2, font * 0.25)
  const s = Math.min(horizontal ? h : w, font * 1.2) * 0.25
  const d = horizontal
    ? `M${num(x + w - font * 1.6 - s)} ${num(y - 1)}L${num(x + w - font * 1.6 + s)} ${num(y + h + 1)}L${num(x + w - font * 1.6 + s + t)} ${num(y + h + 1)}L${num(x + w - font * 1.6 - s + t)} ${num(y - 1)}Z`
    : `M${num(x - 1)} ${num(y + font * 1.6 + s)}L${num(x + w + 1)} ${num(y + font * 1.6 - s)}L${num(x + w + 1)} ${num(y + font * 1.6 - s + t)}L${num(x - 1)} ${num(y + font * 1.6 + s + t)}Z`
  return path(d, INK.white)
}
