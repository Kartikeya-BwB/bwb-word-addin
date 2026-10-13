// The chart builder's form state and how it becomes a ChartSpec. Pure, so it is tested without a browser.

import type { Grid } from '../../charts/grid.ts'
import { gridToText } from '../../charts/grid.ts'
import { parsePasted, toSeriesTable } from '../../charts/data.ts'
import type { NumberFormat, Scale } from '../../charts/format.ts'
import { SAMPLES } from '../../charts/samples.ts'
import type { ChartSpec, ChartType, SortOrder, TimeAxis } from '../../charts/spec.ts'
import { DEFAULT_SIZES, defaultOptions } from '../../charts/spec.ts'
import { CONTENT_AREA, GUTTER } from '../../brand/slideGrid.ts'
import type { Host } from '../../host.ts'

export type Units = '' | Exclude<Scale, ''> | '%'
export type SizePreset = 'full' | 'half' | 'custom'

export interface BuilderState {
  type: ChartType
  grid: Grid
  /** True while the grid holds the type's example data, so changing type swaps in the next example. */
  example: boolean
  title: string
  source: string
  sourceInChart: boolean
  labels: boolean
  legend: boolean
  valueAxis: boolean
  markers: boolean
  highlight: string
  sort: SortOrder
  gapWidth: number
  /** '' for automatic, or '0' to '3'. */
  decimals: string
  units: Units
  currency: string
  axisMin: string
  axisMax: string
  /** Donut centre; empty for the total. */
  centreLabel: string
  today: boolean
  timeAxis: TimeAxis
  seriesOrder: number[]
  /** Palette index for each data series, by position in the data. */
  colours: Array<number | null>
  size: SizePreset
  /** Centimetres in Word, points in PowerPoint. */
  customWidth: string
  customHeight: string
}

/** Which options a chart type uses, so the form shows only those. */
export const OPTION_TYPES = {
  format: ['column', 'stackedColumn', 'bar', 'stackedBar', 'line', 'waterfall', 'mekko', 'donut'],
  highlight: ['column', 'bar', 'donut', 'harvey'],
  labels: ['column', 'stackedColumn', 'column100', 'bar', 'stackedBar', 'bar100', 'line', 'waterfall', 'mekko', 'gantt', 'donut'],
  legend: ['column', 'stackedColumn', 'column100', 'bar', 'stackedBar', 'bar100', 'line', 'waterfall', 'mekko', 'gantt', 'donut'],
  valueAxis: ['column', 'stackedColumn', 'column100', 'bar', 'stackedBar', 'bar100', 'line', 'waterfall', 'mekko'],
  axisRange: ['column', 'stackedColumn', 'bar', 'stackedBar', 'line', 'waterfall'],
  sort: ['column', 'stackedColumn', 'column100', 'bar', 'stackedBar', 'bar100', 'mekko', 'donut'],
  gap: ['column', 'stackedColumn', 'column100', 'bar', 'stackedBar', 'bar100', 'waterfall'],
  series: ['column', 'stackedColumn', 'column100', 'bar', 'stackedBar', 'bar100', 'line', 'mekko'],
  markers: ['line'],
  centre: ['donut'],
  time: ['gantt'],
} as const satisfies Record<string, readonly ChartType[]>

export type OptionName = keyof typeof OPTION_TYPES

export const uses = (type: ChartType, option: OptionName): boolean => (OPTION_TYPES[option] as readonly ChartType[]).includes(type)

const unitsFrom = (hint: NumberFormat): Units => (hint.percent ? '%' : (hint.scale ?? ''))

/**
 * Units and currency read from new data, so pasted "USD 1.2bn" sets USD and bn. Plain numbers keep the author's own
 * units, unless they replace the example. Series order and colours start afresh when the series change.
 */
function fromData(state: BuilderState, grid: Grid): Partial<BuilderState> {
  const table = toSeriesTable(grid)
  const hint = table.hint
  const hinted = Boolean(hint.currency || hint.scale || hint.percent)
  const sameSeries = table.series.length === toSeriesTable(state.grid).series.length
  return {
    ...(hinted || state.example ? { units: unitsFrom(hint), currency: hint.currency ?? '' } : {}),
    ...(sameSeries ? {} : { seriesOrder: [], colours: [] }),
  }
}

function typeDefaults(type: ChartType): Pick<BuilderState, 'labels' | 'legend' | 'valueAxis' | 'markers' | 'gapWidth' | 'sort'> {
  const d = defaultOptions(type)
  return {
    labels: d.labels ?? true,
    legend: d.legend ?? true,
    valueAxis: d.valueAxis ?? false,
    markers: d.markers ?? true,
    gapWidth: d.gapWidth ?? 60,
    sort: d.sort ?? 'none',
  }
}

function sampleFields(type: ChartType): Pick<BuilderState, 'grid' | 'example' | 'title' | 'source' | 'highlight' | 'units' | 'currency'> {
  const sample = SAMPLES[type]
  const hint = { ...toSeriesTable(sample.data).hint, ...sample.format }
  return {
    grid: sample.data.map((r) => [...r]),
    example: true,
    title: sample.title,
    source: sample.source,
    highlight: sample.highlight ?? '',
    units: unitsFrom(hint),
    currency: hint.currency ?? '',
  }
}

export function initialState(host: Host, type: ChartType = 'column'): BuilderState {
  return {
    type,
    ...sampleFields(type),
    ...typeDefaults(type),
    sourceInChart: host === 'powerpoint',
    decimals: '',
    axisMin: '',
    axisMax: '',
    centreLabel: '',
    today: false,
    timeAxis: 'auto',
    seriesOrder: [],
    colours: [],
    size: 'full',
    customWidth: host === 'powerpoint' ? '880' : '16.6',
    customHeight: host === 'powerpoint' ? '360' : '9',
  }
}

/** A new chart type: example data follows the type until the author types or pastes their own. */
export function withType(state: BuilderState, type: ChartType): BuilderState {
  const next = { ...state, type, ...typeDefaults(type) }
  return state.example ? { ...next, ...sampleFields(type), seriesOrder: [], colours: [] } : next
}

export function withExample(state: BuilderState): BuilderState {
  return { ...state, ...sampleFields(state.type), seriesOrder: [], colours: [] }
}

/** The example's title, source and highlight go when the author's own data replaces it. */
const leaveExample = (state: BuilderState): Partial<BuilderState> => (state.example ? { title: '', source: '', highlight: '' } : {})

/** The author's own data from the grid. */
export function withGrid(state: BuilderState, grid: Grid): BuilderState {
  return { ...state, ...leaveExample(state), grid, example: false }
}

/** Pasted text: replaces the grid and takes the units the cells carry. */
export function withPaste(state: BuilderState, text: string): BuilderState {
  const { grid } = parsePasted(text)
  return { ...state, ...leaveExample(state), ...fromData(state, grid), grid, example: false }
}

export const textOf = (state: BuilderState): string => gridToText(state.grid)

export interface Placement {
  widthPt: number
  heightPt: number
  fontPt: number
  /** PowerPoint: top-left corner on the slide. */
  left?: number
  top?: number
}

const CM = 72 / 2.54
const HALF_SLIDE = (CONTENT_AREA.width - GUTTER) / 2

export function sizeOptions(host: Host): Array<{ value: SizePreset; label: string }> {
  return host === 'powerpoint'
    ? [
        { value: 'full', label: 'Content area, 880 \u{00D7} 360 pt' },
        { value: 'half', label: `Half the content area, ${HALF_SLIDE} \u{00D7} 360 pt` },
        { value: 'custom', label: 'Custom size' },
      ]
    : [
        { value: 'full', label: 'Full width, 16.6 \u{00D7} 9 cm' },
        { value: 'half', label: 'Half width, 8.2 \u{00D7} 9 cm' },
        { value: 'custom', label: 'Custom size' },
      ]
}

const number = (text: string, fallback: number, min: number, max: number): number => {
  const n = Number(text.replace(',', '.'))
  return Number.isFinite(n) && n > 0 ? Math.min(max, Math.max(min, n)) : fallback
}

const round2 = (v: number): number => Math.round(v * 100) / 100

/** Size, label size and, in PowerPoint, position for the chosen preset. */
export function placement(state: BuilderState, host: Host): Placement {
  if (host === 'powerpoint') {
    const at = { left: CONTENT_AREA.left, top: CONTENT_AREA.top }
    if (state.size === 'full') return { ...DEFAULT_SIZES.powerpoint, fontPt: 12, ...at }
    if (state.size === 'half') return { widthPt: HALF_SLIDE, heightPt: CONTENT_AREA.height, fontPt: 11, ...at }
    const widthPt = number(state.customWidth, 880, 72, 960)
    const heightPt = number(state.customHeight, 360, 54, 540)
    return {
      widthPt, heightPt, fontPt: widthPt >= 600 ? 12 : widthPt >= 400 ? 11 : 10,
      left: round2(Math.max(0, CONTENT_AREA.left + (CONTENT_AREA.width - widthPt) / 2)),
      top: round2(Math.max(0, CONTENT_AREA.top + (CONTENT_AREA.height - heightPt) / 2)),
    }
  }
  if (state.size === 'full') return { ...DEFAULT_SIZES.word, fontPt: 9 }
  if (state.size === 'half') return { widthPt: round2(8.2 * CM), heightPt: round2(9 * CM), fontPt: 8 }
  const widthPt = round2(number(state.customWidth, 16.6, 2.5, 16.6) * CM)
  const heightPt = round2(number(state.customHeight, 9, 2, 24) * CM)
  return { widthPt, heightPt, fontPt: widthPt >= 400 ? 9 : 8 }
}

function formatOf(state: BuilderState): NumberFormat {
  const decimals = state.decimals === '' ? undefined : Number(state.decimals)
  const scale: Scale = state.units === '%' ? '' : state.units
  return {
    currency: state.currency.trim().toUpperCase().slice(0, 3),
    scale,
    percent: state.units === '%',
    ...(decimals !== undefined && Number.isFinite(decimals) ? { decimals } : {}),
  }
}

const optionalNumber = (text: string): number | undefined => {
  if (!text.trim()) return undefined
  const n = Number(text.replace(/,/g, ''))
  return Number.isFinite(n) ? n : undefined
}

/** The spec the preview draws and the Insert button inserts. `today` is an ISO date for the Gantt today line. */
export function toSpec(state: BuilderState, host: Host, today: string): ChartSpec {
  const place = placement(state, host)
  const min = optionalNumber(state.axisMin)
  const max = optionalNumber(state.axisMax)
  const axis = uses(state.type, 'axisRange') && (min !== undefined || max !== undefined)
    ? { axis: { ...(min !== undefined ? { min } : {}), ...(max !== undefined ? { max } : {}) } }
    : {}
  return {
    type: state.type,
    data: state.grid,
    size: { widthPt: place.widthPt, heightPt: place.heightPt },
    fontPt: place.fontPt,
    format: formatOf(state),
    labels: state.labels,
    legend: state.legend,
    valueAxis: state.valueAxis,
    markers: state.markers,
    gapWidth: state.gapWidth,
    sort: uses(state.type, 'sort') ? state.sort : 'none',
    timeAxis: state.timeAxis,
    sourceInChart: state.sourceInChart,
    ...(state.highlight && uses(state.type, 'highlight') ? { highlight: state.highlight } : {}),
    ...(state.seriesOrder.length ? { seriesOrder: state.seriesOrder } : {}),
    ...(state.colours.length ? { colours: state.colours } : {}),
    ...(state.centreLabel.trim() ? { centreLabel: state.centreLabel.trim() } : {}),
    ...(state.today ? { today } : {}),
    ...(state.title.trim() ? { title: state.title.trim() } : {}),
    ...(state.source.trim() ? { source: state.source.trim() } : {}),
    ...axis,
  }
}
