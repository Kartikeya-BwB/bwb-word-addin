// Alt text that says what a chart shows, for screen readers and for the picture's description in Word and
// PowerPoint: "Column chart: Funding by year, 2022 to 2025. 2022: USD 120m; 2023: USD 150m; ..."

import { chartTypeInfo } from './catalogue.ts'
import { toSeriesTable } from './data.ts'
import { formatDay, parseDate } from './dates.ts'
import type { NumberFormat } from './format.ts'
import { autoDecimals, formatNumber } from './format.ts'
import type { ChartSpec } from './spec.ts'
import { orderedIndices } from './series.ts'
import { ganttRows } from './types/gantt.ts'
import { harveyTable } from './types/harvey.ts'
import { waterfallSteps } from './types/waterfall.ts'

const MAX_ITEMS = 12

/** "A, B and C", the UK way (no comma before "and"). */
export function listOf(items: string[]): string {
  if (items.length <= 1) return items[0] ?? ''
  return `${items.slice(0, -1).join(', ')} and ${items[items.length - 1] ?? ''}`
}

/** True when categories read as a sequence of periods (years, quarters, months or dates), so "2022 to 2025" makes sense. */
export function isSequence(categories: string[]): boolean {
  return (
    categories.length > 1 &&
    categories.every((c) => /^(?:FY|CY)?\s?'?\d{2,4}[A-Z]?$/i.test(c.trim()) || /^H[12]\s/i.test(c.trim()) || parseDate(c) !== null)
  )
}

const clean = (text: string | undefined): string => (text ?? '').trim().replace(/[.\s]+$/, '')

function seriesDetail(spec: ChartSpec, format: NumberFormat): { range: string; detail: string } {
  const table = toSeriesTable(spec.data)
  const categories = table.categories
  const range = isSequence(categories) ? `${categories[0] ?? ''} to ${categories[categories.length - 1] ?? ''}` : ''
  const order = orderedIndices(table.series.length, spec.seriesOrder)
  const series = order.map((i) => table.series[i]).filter((s) => s !== undefined)
  const values = series.flatMap((s) => s.values.filter((v): v is number => v !== null))
  const decimals = format.decimals ?? autoDecimals(values)
  const write = (v: number): string => formatNumber(v, format, { decimals })
  const percentType = spec.type === 'column100' || spec.type === 'bar100' || spec.type === 'mekko'
  const first = series[0]
  if (spec.type === 'donut' && first) {
    const total = first.values.reduce<number>((sum, v) => sum + Math.max(0, v ?? 0), 0)
    const parts = categories.slice(0, MAX_ITEMS).map((c, i) => {
      const v = Math.max(0, first.values[i] ?? 0)
      const share = total > 0 ? formatNumber((v / total) * 100, { percent: true }, { decimals: 0 }) : ''
      return format.percent ? `${c} ${write(v)}` : `${c} ${write(v)} (${share})`
    })
    return { range, detail: `${parts.join(', ')}.` }
  }
  if (series.length === 1 && first && !percentType && categories.length <= MAX_ITEMS) {
    const parts = categories.map((c, i) => {
      const v = first.values[i]
      return `${c}: ${v === null || v === undefined ? 'no value' : write(v)}`
    })
    return { range, detail: `${parts.join('; ')}.` }
  }
  const names = series.map((s) => s.name)
  const seriesText = names.length > 1 ? `Series: ${listOf(names.slice(0, MAX_ITEMS))}.` : ''
  const categoryText = range || categories.length > MAX_ITEMS ? `${categories.length} categories.` : `Categories: ${listOf(categories)}.`
  return { range, detail: [seriesText, categoryText].filter(Boolean).join(' ') }
}

/** Alt text for a chart: its type and title, the span it covers, then what it shows. */
export function describeChart(spec: ChartSpec, format: NumberFormat): string {
  const info = chartTypeInfo(spec.type)
  const title = clean(spec.title)
  const head = title ? `${info.name}: ${title}` : info.name
  const sentence = (range: string, detail: string): string => `${head}${range ? `, ${range}` : ''}.${detail ? ` ${detail}` : ''}`
  if (spec.type === 'waterfall') {
    const steps = waterfallSteps(spec.data)
    const start = steps[0]
    const end = steps[steps.length - 1]
    if (!start || !end) return sentence('', '')
    const decimals = format.decimals ?? autoDecimals(steps.map((s) => s.value))
    const write = (v: number): string => formatNumber(v, format, { decimals })
    const changes = steps.filter((s) => s.kind === 'delta').length
    return sentence('', `From ${start.label} (${write(start.to)}) to ${end.label} (${write(end.to)}), in ${changes} change${changes === 1 ? '' : 's'}.`)
  }
  if (spec.type === 'gantt') {
    const rows = ganttRows(spec.data).filter((r) => r.kind !== 'group')
    if (!rows.length) return sentence('', '')
    const from = Math.min(...rows.map((r) => r.start))
    const to = Math.max(...rows.map((r) => r.end))
    const tasks = rows.filter((r) => r.kind === 'task').length
    const milestones = rows.length - tasks
    const counts = [`${tasks} task${tasks === 1 ? '' : 's'}`, ...(milestones ? [`${milestones} milestone${milestones === 1 ? '' : 's'}`] : [])]
    const range = `${formatDay(from, 'month')} to ${formatDay(to, 'month')}`
    return sentence(range, `${listOf(counts)}: ${listOf(rows.slice(0, MAX_ITEMS).map((r) => r.label))}.`)
  }
  if (spec.type === 'harvey') {
    const table = harveyTable(spec.data)
    const n = table.options.length
    const m = table.criteria.length
    return sentence('', `${n} option${n === 1 ? '' : 's'} scored from 0 to 4 against ${m} criteri${m === 1 ? 'on' : 'a'}: ${listOf(table.criteria)}.`)
  }
  const { range, detail } = seriesDetail(spec, format)
  return sentence(range, detail)
}
