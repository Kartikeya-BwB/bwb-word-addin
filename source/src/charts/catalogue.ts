// The chart types the builder offers, with the names alt text uses.

export type ChartType =
  | 'column'
  | 'stackedColumn'
  | 'column100'
  | 'bar'
  | 'stackedBar'
  | 'bar100'
  | 'line'
  | 'area'
  | 'area100'
  | 'combo'
  | 'waterfall'
  | 'mekko'
  | 'scatter'
  | 'bubble'
  | 'gantt'
  | 'donut'
  | 'harvey'
  | 'map'

export interface ChartTypeInfo {
  type: ChartType
  /** Tile label. */
  label: string
  /** How alt text names it. */
  name: string
  hint: string
}

export const CHART_TYPES: readonly ChartTypeInfo[] = [
  { type: 'column', label: 'Column', name: 'Column chart', hint: 'Compare values across a few categories or years' },
  { type: 'stackedColumn', label: 'Stacked column', name: 'Stacked column chart', hint: 'Totals and what makes them up' },
  { type: 'column100', label: '100% column', name: '100% stacked column chart', hint: 'Shares within each category' },
  { type: 'bar', label: 'Bar', name: 'Bar chart', hint: 'Rank categories with long names' },
  { type: 'stackedBar', label: 'Stacked bar', name: 'Stacked bar chart', hint: 'Ranked totals and their parts' },
  { type: 'bar100', label: '100% bar', name: '100% stacked bar chart', hint: 'Shares, such as survey answers' },
  { type: 'line', label: 'Line', name: 'Line chart', hint: 'Trends over time' },
  { type: 'area', label: 'Stacked area', name: 'Stacked area chart', hint: 'A total over time and what makes it up' },
  { type: 'area100', label: '100% area', name: '100% area chart', hint: 'Shares over time' },
  { type: 'combo', label: 'Combination', name: 'Combination chart', hint: 'Columns with one or two lines, on a second axis if needed' },
  { type: 'waterfall', label: 'Waterfall', name: 'Waterfall chart', hint: 'How a start value becomes an end value' },
  { type: 'mekko', label: 'Mekko', name: 'Mekko chart', hint: 'Market size and mix: widths are totals' },
  { type: 'scatter', label: 'Scatter', name: 'Scatter chart', hint: 'Two measures for each item, with quadrants for a 2x2' },
  { type: 'bubble', label: 'Bubble', name: 'Bubble chart', hint: 'A scatter with a third measure as the bubble size' },
  { type: 'gantt', label: 'Gantt', name: 'Gantt chart', hint: 'Workplan with tasks and milestones' },
  { type: 'donut', label: 'Donut', name: 'Donut chart', hint: 'Parts of one whole, six at most' },
  { type: 'harvey', label: 'Harvey balls', name: 'Harvey ball table', hint: 'Options scored 0 to 4 against criteria' },
  { type: 'map', label: 'Data map', name: 'Data map', hint: 'Values or groups by country, shaded in BwB greens' },
]

export const chartTypeInfo = (type: ChartType): ChartTypeInfo =>
  CHART_TYPES.find((t) => t.type === type) ?? { type, label: type, name: 'Chart', hint: '' }
