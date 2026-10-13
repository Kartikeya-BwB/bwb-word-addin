// Series as they are drawn: in the chosen order, with their palette colours, categories sorted when asked, and one
// category in green with the rest in navy when a highlight is set.

import { COLOURS } from '../brand/tokens.ts'
import type { SeriesTable } from './data.ts'
import { toSeriesTable } from './data.ts'
import type { ChartContext } from './frame.ts'
import { warn } from './frame.ts'
import { paletteColour } from './palette.ts'

export const MAX_SERIES = 6

export interface DrawSeries {
  name: string
  /** Position of the series in the data, which its colour follows. */
  index: number
  colour: string
  values: Array<number | null>
}

export interface Prepared {
  categories: string[]
  series: DrawSeries[]
  table: SeriesTable
  /** Colour of one data point, which differs from its series colour only for a highlighted category. */
  colourAt: (series: DrawSeries, category: number) => string
  /** Index of the highlighted category, or -1. */
  highlighted: number
}

const same = (a: string, b: string): boolean => a.trim().toLowerCase() === b.trim().toLowerCase()

/** The colour of a data series: the palette entry the spec gives it, or the next one in BwB order. */
export function seriesColour(ctx: ChartContext, index: number): string {
  const chosen = ctx.spec.colours?.[index]
  return paletteColour(typeof chosen === 'number' && chosen >= 0 ? chosen : index)
}

/** Data series in the spec's order; any series the order leaves out follow in their own order. */
export function orderedIndices(count: number, order: readonly number[] | undefined): number[] {
  const seen = new Set<number>()
  const out: number[] = []
  for (const i of [...(order ?? []), ...Array.from({ length: count }, (_, k) => k)]) {
    if (Number.isInteger(i) && i >= 0 && i < count && !seen.has(i)) {
      seen.add(i)
      out.push(i)
    }
  }
  return out
}

export function prepareSeries(ctx: ChartContext, options: { sortable: boolean; highlight: boolean }): Prepared {
  const table = toSeriesTable(ctx.spec.data)
  if (table.invalid.length) {
    const n = table.invalid.length
    warn(ctx, `${n} cell${n === 1 ? ' is' : 's are'} not a number and ${n === 1 ? 'is' : 'are'} left as a gap.`)
  }
  let series: DrawSeries[] = orderedIndices(table.series.length, ctx.spec.seriesOrder).map((index) => ({
    name: table.series[index]?.name ?? '',
    index,
    colour: seriesColour(ctx, index),
    values: table.series[index]?.values ?? [],
  }))
  if (series.length > MAX_SERIES) {
    warn(ctx, `Charts show up to ${MAX_SERIES} series; the rest are left out.`)
    series = series.slice(0, MAX_SERIES)
  }
  let order = table.categories.map((_, i) => i)
  const sort = ctx.spec.sort ?? 'none'
  if (options.sortable && sort !== 'none') {
    const total = (c: number): number => series.reduce((sum, s) => sum + (s.values[c] ?? 0), 0)
    order = [...order].sort((a, b) => (sort === 'desc' ? total(b) - total(a) : total(a) - total(b)))
  }
  const categories = order.map((i) => table.categories[i] ?? '')
  series = series.map((s) => ({ ...s, values: order.map((i) => s.values[i] ?? null) }))
  const wanted = ctx.spec.highlight?.trim() ?? ''
  const highlighted = options.highlight && wanted ? categories.findIndex((c) => same(c, wanted)) : -1
  if (wanted && options.highlight && highlighted < 0) warn(ctx, `There is no category called \u{201C}${wanted}\u{201D} to highlight.`)
  const colourAt = (s: DrawSeries, c: number): string => {
    if (highlighted < 0 || series.length > 1) return s.colour
    return c === highlighted ? COLOURS.green : COLOURS.navy
  }
  return { categories, series, table, colourAt, highlighted }
}

/** Every value of every series, for scales and label decimals. */
export const allValues = (series: DrawSeries[]): number[] =>
  series.flatMap((s) => s.values.filter((v): v is number => v !== null && Number.isFinite(v)))
