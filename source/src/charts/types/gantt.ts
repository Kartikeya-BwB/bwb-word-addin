// Gantt charts: one row per workstream or task with a navy bar from its start to its end, milestones as diamonds,
// and a time axis in months, quarters or years across the top. A row with no dates is a group heading for the rows
// under it; a row with one date is a milestone. Rows are separated by horizontal rules only.

import { COLOURS } from '../../brand/tokens.ts'
import type { Period, Precision, TimeUnit } from '../dates.ts'
import { MONTH_NAMES, chooseUnit, civil, formatDay, parseDate, periods } from '../dates.ts'
import type { ChartContext, LegendItem } from '../frame.ts'
import { LABELS_CUT, chartFrame, warn } from '../frame.ts'
import type { Grid } from '../grid.ts'
import { normaliseGrid } from '../grid.ts'
import { fitLabels, lineHeight, textWidth } from '../measure.ts'
import { INK } from '../palette.ts'
import { centredBaseline, diamond, group, line, rect, text, textLinesCentred } from '../svg.ts'
import { placeholder } from './empty.ts'

export interface GanttRow {
  label: string
  kind: 'task' | 'milestone' | 'group'
  /** First day; for a milestone, its date. */
  start: number
  /** Last day, inclusive. */
  end: number
  /** Under a group heading. */
  indent: boolean
  /** How exactly a milestone's date was given; a month or quarter puts it at the end of that period. */
  precision?: Precision
}

/** Reads rows of label, start and end (first row headings). */
export function ganttRows(input: Grid, onWarning: (message: string) => void = () => {}): GanttRow[] {
  const rows: GanttRow[] = []
  let inGroup = false
  for (const row of normaliseGrid(input).slice(1)) {
    const label = (row[0] ?? '').trim()
    const startText = (row[1] ?? '').trim()
    const endText = (row[2] ?? '').trim()
    if (!label && !startText && !endText) continue
    const start = startText ? parseDate(startText) : null
    const end = endText ? parseDate(endText) : null
    if ((startText && !start) || (endText && !end)) {
      onWarning(`\u{201C}${startText && !start ? startText : endText}\u{201D} is not a date; write dates like 5 October 2026, 2026-10-05 or Q4 2026.`)
      continue
    }
    if (!start && !end) {
      rows.push({ label, kind: 'group', start: 0, end: 0, indent: false })
      inGroup = true
      continue
    }
    if (start && end) {
      const [a, b] = start.first <= end.last ? [start.first, end.last] : [end.first, start.last]
      rows.push({ label, kind: 'task', start: a, end: b, indent: inGroup })
      continue
    }
    const one = start ?? end
    if (one) {
      const day = one.precision === 'day' ? one.first : one.last
      rows.push({ label, kind: 'milestone', start: day, end: day, indent: inGroup, precision: one.precision })
    }
  }
  return rows
}

/** A milestone's date as precisely as it was given: "30 Nov", "Mar 2027", "Q1 2027". */
export function milestoneLabel(row: GanttRow, withYear: boolean): string {
  const { year, month } = civil(row.start)
  switch (row.precision ?? 'day') {
    case 'month':
      return `${(MONTH_NAMES[month] ?? '').slice(0, 3)} ${year}`
    case 'quarter':
      return `Q${Math.floor(month / 3) + 1} ${year}`
    case 'half':
      return `H${month < 6 ? 1 : 2} ${year}`
    case 'year':
      return String(year)
    default:
      return withYear ? `${formatDay(row.start, 'short')} ${year}` : formatDay(row.start, 'short')
  }
}

/** The label variant every period can show: "Jan" or "J" for months, or none when even that does not fit. */
function periodLabels(list: Period[], widthOf: (p: Period) => number, size: number): number {
  const variants = list[0]?.labels.length ?? 0
  for (let v = 0; v < variants; v += 1) {
    if (list.every((p) => textWidth(p.labels[v] ?? '', size) <= widthOf(p) - size * 0.3)) return v
  }
  return -1
}

export function renderGantt(ctx: ChartContext): string {
  const rows = ganttRows(ctx.spec.data, (m) => warn(ctx, m))
  const dated = rows.filter((r) => r.kind !== 'group')
  if (!dated.length) return placeholder(ctx, 'Add tasks with start and end dates to draw the plan.')
  const font = ctx.font
  const lh = lineHeight(font)
  const items: LegendItem[] = []
  if (dated.some((r) => r.kind === 'task')) items.push({ label: 'Activity', colour: COLOURS.navy })
  if (dated.some((r) => r.kind === 'milestone')) items.push({ label: 'Milestone', colour: COLOURS.green, mark: 'diamond' })
  const frame = chartFrame(ctx, items)
  const from = Math.min(...dated.map((r) => r.start))
  const to = Math.max(...dated.map((r) => r.end))
  const chosen = ctx.spec.timeAxis && ctx.spec.timeAxis !== 'auto' ? ctx.spec.timeAxis : chooseUnit(from, to)
  const unit: TimeUnit = periods(from, to, chosen).length > 120 ? 'year' : chosen
  const list = periods(from, to, unit)
  const first = list[0]
  const last = list[list.length - 1]
  if (!first || !last) return placeholder(ctx, 'Add tasks with start and end dates to draw the plan.')
  const today = ctx.spec.today ? parseDate(ctx.spec.today)?.first : undefined
  const showToday = today !== undefined && today >= first.start && today < last.end
  const indent = font * 0.9
  const labelWidth = Math.min(
    frame.plot.w * (frame.plot.w < 400 ? 0.4 : 0.3),
    Math.max(...rows.map((r) => textWidth(r.label, font, r.kind === 'group') + (r.indent ? indent : 0))) + font * 0.8,
  )
  const pad = Math.min(lh * 0.8, frame.plot.h * 0.05)
  const timeX = frame.plot.x + labelWidth + (dated.some((r) => r.kind === 'milestone' && r.start - first.start < 7) ? pad : 0)
  const timeW = frame.plot.x + frame.plot.w - timeX - (dated.some((r) => r.kind === 'milestone' && last.end - r.start < 7) ? pad : 0)
  const x = (day: number): number => timeX + ((day - first.start) / (last.end - first.start)) * timeW
  const headerRows = unit === 'year' ? 1 : 2
  const headerH = headerRows * lh + font * 0.35
  const todayH = showToday ? font * 1.3 : 0
  const top = frame.plot.y + headerH
  const rowH = (frame.plot.h - headerH - todayH) / rows.length
  const maxLines = Math.min(2, Math.max(1, Math.floor((rowH * 0.95) / lh)))
  const fitted = fitLabels(rows.map((r) => r.label), rows.map((r) => labelWidth - font * 0.8 - (r.indent ? indent : 0)), font, maxLines)
  if (fitted.wrapped.some((w) => w.truncated)) warn(ctx, LABELS_CUT)
  const withYear = last.end - first.start > 366
  const back: string[] = []
  const shapes: string[] = []
  const labels: string[] = []
  // Header: years above, then months or quarters.
  const unitRow = frame.plot.y + (headerRows - 1) * lh + lh / 2
  if (unit !== 'year') {
    const variant = periodLabels(list, (p) => x(p.end) - x(p.start), font)
    if (variant >= 0) for (const p of list) labels.push(text((x(p.start) + x(p.end)) / 2, centredBaseline(unitRow, font), p.labels[variant] ?? '', { size: font, fill: INK.text, anchor: 'middle' }))
    const years = [...new Set(list.map((p) => p.year))]
    let freeFrom = timeX
    for (const year of years) {
      const inYear = list.filter((p) => p.year === year)
      const a = x(inYear[0]?.start ?? first.start)
      const b = x(inYear[inYear.length - 1]?.end ?? last.end)
      const label = String(year)
      const width = textWidth(label, font, true)
      if (a > timeX) back.push(line(a, frame.plot.y, a, top, INK.grid, 0.75))
      const at = b - a >= width + font * 0.6 ? a + font * 0.3 : null
      if (at !== null && at >= freeFrom) {
        labels.push(text(at, centredBaseline(frame.plot.y + lh / 2, font), label, { size: font, fill: INK.text, bold: true }))
        freeFrom = at + width + font * 0.5
      }
    }
  } else {
    const variant = periodLabels(list, (p) => x(p.end) - x(p.start), font)
    if (variant >= 0) for (const p of list) labels.push(text((x(p.start) + x(p.end)) / 2, centredBaseline(unitRow, font), p.labels[variant] ?? '', { size: font, fill: INK.text, anchor: 'middle', bold: true }))
  }
  back.push(line(frame.plot.x, top, frame.plot.x + frame.plot.w, top, INK.text, 0.75))
  const barH = Math.min(rowH * 0.5, font * 1.3)
  const half = Math.min(rowH * 0.36, font * 0.75)
  rows.forEach((r, i) => {
    const cy = top + rowH * (i + 0.5)
    if (i < rows.length - 1) back.push(line(frame.plot.x, top + rowH * (i + 1), frame.plot.x + frame.plot.w, top + rowH * (i + 1), INK.grid, 0.5))
    const lines = fitted.wrapped[i]?.lines ?? []
    labels.push(textLinesCentred(frame.plot.x + (r.indent ? indent : 0), cy, lines, { size: fitted.size, fill: INK.text, bold: r.kind === 'group' }))
    if (r.kind === 'task') shapes.push(rect(x(r.start), cy - barH / 2, Math.max(1.5, x(r.end + 1) - x(r.start)), barH, COLOURS.navy))
    if (r.kind !== 'milestone') return
    const mx = x(r.start + 0.5)
    shapes.push(diamond(mx, cy, half, COLOURS.green))
    if (!ctx.labels) return
    const label = milestoneLabel(r, withYear)
    const width = textWidth(label, font)
    const right = mx + half + font * 0.35
    const fitsRight = right + width <= frame.plot.x + frame.plot.w
    labels.push(text(fitsRight ? right : mx - half - font * 0.35, centredBaseline(cy, font), label, { size: font, fill: INK.data, anchor: fitsRight ? 'start' : 'end' }))
  })
  if (showToday && today !== undefined) {
    const tx = x(today + 0.5)
    const bottom = top + rowH * rows.length
    shapes.push(line(tx, top, tx, bottom, COLOURS.green, 1))
    labels.push(text(tx, bottom + font * 0.95, 'Today', { size: font, fill: INK.text, anchor: 'middle' }))
  }
  return frame.svg + group('grid', back.join('')) + group('bars', shapes.join('')) + group('labels', labels.join(''))
}
