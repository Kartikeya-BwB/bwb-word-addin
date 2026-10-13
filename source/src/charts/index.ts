// The chart builder's engine: data parsing, the chart spec and renderChart. Pure TypeScript, no Office or React.

export type { ChartSize, ChartSpec, ChartType, ChartTypeInfo, RenderedChart, SortOrder, TimeAxis } from './spec.ts'
export { CHART_TYPES, DEFAULT_SIZES, chartTypeInfo, createContext, defaultOptions, renderChart } from './spec.ts'
export type { Grid, PastedTable, ParsedNumber, SeriesTable } from './data.ts'
export {
  addColumn, addRow, cellAt, gridToText, moveColumn, normaliseGrid, parseNumber, parsePasted, pasteBlock, removeColumn,
  removeRow, setCell, toSeriesTable, transpose, trimGrid, unitsHint,
} from './data.ts'
export type { NumberFormat, Scale } from './format.ts'
export { SCALES, formatNumber, unitsLabel } from './format.ts'
export { PALETTE_NAMES, SERIES_HEX } from './palette.ts'
export { SAMPLES } from './samples.ts'
export type { Sample } from './samples.ts'
export { londonToday, isoDay, parseDate } from './dates.ts'
