// Waterfall charts: a start value, the changes that lead from it, subtotals, and the end value. Each row of the data
// is a step: label, value, and optionally its kind (delta, subtotal or total). A total or subtotal with no value is
// the running total. Increases are green, decreases charcoal, totals navy, with thin connectors between bars.

import { COLOURS } from '../../brand/tokens.ts'
import { columnAxes } from '../cartesian.ts'
import { parseNumber } from '../data.ts'
import { autoDecimals } from '../format.ts'
import type { ChartContext, LegendItem } from '../frame.ts'
import { chartFrame, dropLabels, fmt, unitsNote, warn } from '../frame.ts'
import type { Grid } from '../grid.ts'
import { normaliseGrid } from '../grid.ts'
import { CAP_HEIGHT, MIN_FONT_PT, lineHeight, textWidth } from '../measure.ts'
import { INK } from '../palette.ts'
import { group, line, rect, text } from '../svg.ts'
import { maxBar } from './bars.ts'
import { placeholder } from './empty.ts'

export type StepKind = 'delta' | 'subtotal' | 'total'

export interface Step {
  label: string
  kind: StepKind
  /** The change for a delta; the level for a total or subtotal. */
  value: number
  /** Where the bar starts and ends: from the running total to the new running total, or from zero for totals. */
  from: number
  to: number
}

export const STEP_COLOURS = { increase: COLOURS.green, decrease: COLOURS.charcoal, total: COLOURS.navy } as const

function kindOf(text: string): StepKind | null {
  const t = text.trim().toLowerCase()
  if (!t || /^(?:d|delta|change|step|increase|decrease|\+|-)$/.test(t)) return t ? 'delta' : null
  if (/^(?:s|st|sub|subtotal|sub-total|sub total)$/.test(t)) return 'subtotal'
  if (/^(?:t|e|=|total|start|end|sum|opening|closing|base)$/.test(t)) return 'total'
  return null
}

/**
 * Reads waterfall steps from a grid (first row headings). Without a kind column, the first row is the starting
 * total and any row without a value is a total of everything before it.
 */
export function waterfallSteps(input: Grid, onWarning: (message: string) => void = () => {}): Step[] {
  const grid = normaliseGrid(input)
  const rows = grid.slice(1).filter((r) => r.some((c) => c.trim()))
  const hasKinds = rows.some((r) => kindOf(r[2] ?? '') !== null)
  let running = 0
  const steps: Step[] = []
  rows.forEach((row, i) => {
    const label = (row[0] ?? '').trim() || `Step ${i + 1}`
    const parsed = parseNumber(row[1] ?? '')
    if (!parsed.ok) {
      onWarning(`\u{201C}${row[1] ?? ''}\u{201D} is not a number; that step is left out.`)
      return
    }
    const given = parsed.value
    let kind: StepKind = kindOf(row[2] ?? '') ?? 'delta'
    if (!hasKinds) kind = i === 0 || given === null ? 'total' : 'delta'
    else if (i === 0 && kindOf(row[2] ?? '') === null) kind = 'total'
    if (kind === 'delta') {
      if (given === null) return
      steps.push({ label, kind, value: given, from: running, to: running + given })
      running += given
      return
    }
    const level = given ?? running
    steps.push({ label, kind, value: level, from: 0, to: level })
    running = level
  })
  return steps
}

export function stepColour(step: Step): string {
  if (step.kind !== 'delta') return STEP_COLOURS.total
  return step.value >= 0 ? STEP_COLOURS.increase : STEP_COLOURS.decrease
}

export function renderWaterfall(ctx: ChartContext): string {
  const steps = waterfallSteps(ctx.spec.data, (m) => warn(ctx, m))
  if (!steps.length) return placeholder(ctx, 'Add steps with values to draw the waterfall.')
  const font = ctx.font
  const decimals = ctx.format.decimals ?? autoDecimals(steps.map((s) => s.value))
  const bare = Boolean(unitsNote(ctx.format)) && ctx.valueAxis
  const write = (s: Step): string => fmt(ctx, s.value, { decimals, bare, plus: s.kind === 'delta' })
  const present = new Set(steps.map(stepColour))
  const items: LegendItem[] = [
    { label: 'Increase', colour: STEP_COLOURS.increase },
    { label: 'Decrease', colour: STEP_COLOURS.decrease },
    { label: 'Total', colour: STEP_COLOURS.total },
  ].filter((item) => present.has(item.colour))
  const frame = chartFrame(ctx, items, bare ? unitsNote(ctx.format) : '')
  const levels = steps.flatMap((s) => [s.from, s.to])
  const lo = Math.min(0, ...levels)
  const hi = Math.max(0, ...levels)
  // A label under a falling bar needs room below the plot only when that bar ends near the bottom of the scale.
  const lowest = Math.min(...steps.filter((s) => s.to < s.from || s.to < 0).map((s) => s.to), Number.POSITIVE_INFINITY)
  const nearBottom = ((lowest - lo) / (hi - lo || 1)) * frame.plot.h * 0.7 < lineHeight(font) * 1.4
  const axes = columnAxes(ctx, frame.plot, steps.map((s) => s.label), { lo, hi, roomAbove: ctx.labels, roomBelow: ctx.labels && nearBottom })
  const barWidth = Math.min(axes.band / (1 + ctx.gap), maxBar(ctx, 1))
  const texts = steps.map(write)
  let size: number | null = null
  for (let t = font; t >= MIN_FONT_PT && size === null; t -= 0.5) if (texts.every((tx) => textWidth(tx, t) <= axes.band - t * 0.3)) size = t
  if (ctx.labels && size === null) dropLabels(ctx)
  const bars: string[] = []
  const connectors: string[] = []
  const labels: string[] = []
  steps.forEach((s, i) => {
    const x = axes.bandStart(i) + (axes.band - barWidth) / 2
    const yFrom = axes.at(s.from)
    const yTo = axes.at(s.to)
    bars.push(rect(x, Math.min(yFrom, yTo), barWidth, Math.max(0.75, Math.abs(yTo - yFrom)), stepColour(s)))
    const next = steps[i + 1]
    if (next) {
      const nextX = axes.bandStart(i + 1) + (axes.band - barWidth) / 2
      connectors.push(line(x + barWidth, yTo, nextX, yTo, INK.charcoal, 0.5))
    }
    if (!ctx.labels || size === null) return
    const up = s.to >= s.from && !(s.kind !== 'delta' && s.to < 0)
    const y = up ? Math.min(yFrom, yTo) - size * 0.35 : Math.max(yFrom, yTo) + size * (0.35 + CAP_HEIGHT)
    labels.push(text(x + barWidth / 2, y, texts[i] ?? '', { size, fill: INK.data, anchor: 'middle', bold: s.kind !== 'delta' }))
  })
  return frame.svg + axes.back + group('connectors', connectors.join('')) + group('bars', bars.join('')) + axes.front + group('labels', labels.join(''))
}
