// What a chart is: its type, its data as cell text, and its options. renderChart turns a spec into SVG markup in
// points, with alt text. The title and source are for the caption; neither is drawn, except the source line when
// sourceInChart is set.

import { describeChart } from './alt.ts'
import type { ChartType } from './catalogue.ts'
import type { NumberFormat } from './format.ts'
import type { ChartContext } from './frame.ts'
import { LABELS_DROPPED, LABELS_TO_AXIS, defaultFontPt } from './frame.ts'
import type { Grid } from './grid.ts'
import { MIN_FONT_PT } from './measure.ts'
import { toSeriesTable } from './data.ts'
import { svgDocument } from './svg.ts'
import { renderBars, renderColumns } from './types/bars.ts'
import { renderDonut } from './types/donut.ts'
import { renderGantt } from './types/gantt.ts'
import { renderHarvey } from './types/harvey.ts'
import { renderLine } from './types/line.ts'
import { renderMekko } from './types/mekko.ts'
import { renderWaterfall } from './types/waterfall.ts'

export type { ChartType, ChartTypeInfo } from './catalogue.ts'
export { CHART_TYPES, chartTypeInfo } from './catalogue.ts'

export type SortOrder = 'none' | 'desc' | 'asc'
export type TimeAxis = 'auto' | 'month' | 'quarter' | 'year'

export interface ChartSize {
  widthPt: number
  heightPt: number
}

export interface ChartSpec {
  type: ChartType
  /** Cell text: series names in the first row, categories in the first column. */
  data: Grid
  /** Data series indices (0 is the second column) in drawing and legend order. */
  seriesOrder?: number[]
  /** Palette index for each data series, by its position in the data; navy, green, green 50, charcoal, light grey. */
  colours?: Array<number | null>
  /** A category shown in green with the rest in navy (one-series column and bar charts, donuts, Harvey ball rows). */
  highlight?: string
  format?: NumberFormat
  axis?: { min?: number; max?: number }
  /** Data labels; default on. */
  labels?: boolean
  /** Legend at the top left; default on. */
  legend?: boolean
  /** Value axis labels and horizontal gridlines; default on for line charts only. */
  valueAxis?: boolean
  /** Gap between bars as a percentage of the bar width; default 60. */
  gapWidth?: number
  sort?: SortOrder
  /** Line chart markers; default on. */
  markers?: boolean
  /** Text in the middle of a donut; default the total. An empty string leaves it blank. */
  centreLabel?: string
  /** Gantt "today" line at this ISO date. */
  today?: string
  timeAxis?: TimeAxis
  size: ChartSize
  /** Label size; by default 12 pt on slides and 9 pt in documents. */
  fontPt?: number
  /** For the caption and alt text; never drawn. */
  title?: string
  /** For the caption, or drawn at the bottom left when sourceInChart is set. */
  source?: string
  sourceInChart?: boolean
}

export interface RenderedChart {
  svg: string
  widthPt: number
  heightPt: number
  altText: string
  /** Things the author should know, such as labels left out because they did not fit. */
  warnings: string[]
}

const CM = 72 / 2.54

/** Default sizes: the full A4 text width by 9 cm in Word, the slide content area in PowerPoint. */
export const DEFAULT_SIZES: Record<'word' | 'powerpoint', ChartSize> = {
  word: { widthPt: Math.round(16.6 * CM * 100) / 100, heightPt: Math.round(9 * CM * 100) / 100 },
  powerpoint: { widthPt: 880, heightPt: 360 },
}

/** Options that suit each type before the author changes anything. */
export function defaultOptions(type: ChartType): Partial<ChartSpec> {
  return {
    labels: true,
    legend: type !== 'waterfall',
    valueAxis: type === 'line',
    gapWidth: type === 'waterfall' ? 40 : 60,
    markers: true,
    sort: 'none',
  }
}

const clampSize = (v: number, fallback: number): number => (Number.isFinite(v) && v >= 36 ? Math.min(v, 4000) : fallback)

export function createContext(spec: ChartSpec): ChartContext {
  const width = clampSize(spec.size.widthPt, DEFAULT_SIZES.word.widthPt)
  const height = clampSize(spec.size.heightPt, DEFAULT_SIZES.word.heightPt)
  const hint = toSeriesTable(spec.data).hint
  const defaults = defaultOptions(spec.type)
  return {
    spec,
    width,
    height,
    font: Math.max(MIN_FONT_PT, spec.fontPt ?? defaultFontPt(width)),
    labels: spec.labels ?? defaults.labels ?? true,
    legend: spec.legend ?? defaults.legend ?? true,
    valueAxis: spec.valueAxis ?? defaults.valueAxis ?? false,
    gap: Math.min(5, Math.max(0, (spec.gapWidth ?? defaults.gapWidth ?? 60) / 100)),
    format: { ...hint, ...spec.format },
    warnings: [],
    labelsDropped: false,
  }
}

/** Types whose values can be read from a value axis when their labels do not fit. */
const AXIS_FALLBACK = new Set<ChartType>(['column', 'stackedColumn', 'bar', 'stackedBar', 'line', 'waterfall'])

function draw(ctx: ChartContext): string {
  switch (ctx.spec.type) {
    case 'column':
      return renderColumns(ctx, 'clustered')
    case 'stackedColumn':
      return renderColumns(ctx, 'stacked')
    case 'column100':
      return renderColumns(ctx, 'percent')
    case 'bar':
      return renderBars(ctx, 'clustered')
    case 'stackedBar':
      return renderBars(ctx, 'stacked')
    case 'bar100':
      return renderBars(ctx, 'percent')
    case 'line':
      return renderLine(ctx)
    case 'waterfall':
      return renderWaterfall(ctx)
    case 'mekko':
      return renderMekko(ctx)
    case 'gantt':
      return renderGantt(ctx)
    case 'donut':
      return renderDonut(ctx)
    case 'harvey':
      return renderHarvey(ctx)
  }
}

/** Draws a chart: SVG in points, its size, alt text, and any warnings for the author. */
export function renderChart(spec: ChartSpec): RenderedChart {
  let ctx = createContext(spec)
  let body = draw(ctx)
  if (ctx.labelsDropped && !ctx.valueAxis && AXIS_FALLBACK.has(spec.type)) {
    const retry: ChartContext = { ...ctx, valueAxis: true, warnings: [], labelsDropped: false }
    body = draw(retry)
    ctx = { ...retry, warnings: [...retry.warnings.filter((w) => w !== LABELS_DROPPED), LABELS_TO_AXIS] }
  }
  const altText = describeChart(spec, ctx.format)
  return { svg: svgDocument(ctx.width, ctx.height, body, altText), widthPt: ctx.width, heightPt: ctx.height, altText, warnings: ctx.warnings }
}
