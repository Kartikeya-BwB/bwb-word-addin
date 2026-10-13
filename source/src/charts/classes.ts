// Classes for data maps: values grouped into 3 to 5 bands, by quantiles (about as many countries in each band) or
// by equal intervals, with round boundaries so the legend reads "10 to 25" rather than "9.83 to 24.71".

import { niceStep } from './axes.ts'

export type ClassMethod = 'quantile' | 'equal'

/** A boundary rounded to two significant figures: 9,834 becomes 9,800 and 0.0347 becomes 0.035. */
export function roundBoundary(value: number): number {
  if (value === 0 || !Number.isFinite(value)) return 0
  const power = 10 ** (Math.floor(Math.log10(Math.abs(value))) - 1)
  return Number.parseFloat((Math.round(value / power) * power).toPrecision(12))
}

/** The value at a fraction of the way through sorted values, interpolating between neighbours. */
export function quantile(sorted: number[], q: number): number {
  if (!sorted.length) return 0
  const position = (sorted.length - 1) * Math.min(1, Math.max(0, q))
  const below = Math.floor(position)
  const low = sorted[below] ?? 0
  const high = sorted[Math.min(sorted.length - 1, below + 1)] ?? low
  return low + (high - low) * (position - below)
}

/**
 * The boundaries between classes, ascending; a value equal to a boundary belongs to the class above it. Fewer than
 * `count - 1` boundaries come back when the values cannot fill that many classes.
 */
export function classBreaks(values: number[], count: number, method: ClassMethod): number[] {
  const sorted = values.filter(Number.isFinite).sort((a, b) => a - b)
  const k = Math.min(5, Math.max(2, Math.round(count)))
  const min = sorted[0]
  const max = sorted[sorted.length - 1]
  if (min === undefined || max === undefined || min === max) return []
  let raw: number[]
  if (method === 'equal') {
    const step = niceStep((max - min) / k)
    const start = Math.floor(min / step) * step
    raw = Array.from({ length: k - 1 }, (_, i) => start + step * (i + 1))
  } else {
    raw = Array.from({ length: k - 1 }, (_, i) => roundBoundary(quantile(sorted, (i + 1) / k)))
  }
  const breaks: number[] = []
  for (const b of raw.map((v) => Number.parseFloat(v.toPrecision(12)))) {
    if (b > min && b <= max && (breaks.length === 0 || b > (breaks[breaks.length - 1] ?? b))) breaks.push(b)
  }
  return breaks
}

/** The class a value falls in: 0 below the first boundary, up to the number of boundaries. */
export function classOf(value: number, breaks: number[]): number {
  let i = 0
  while (i < breaks.length && value >= (breaks[i] ?? Number.POSITIVE_INFINITY)) i += 1
  return i
}

/** Legend labels: "Under 10", "10 to 25", "25 and over". */
export function classLabels(breaks: number[], write: (value: number) => string): string[] {
  if (!breaks.length) return ['All values']
  return [
    `Under ${write(breaks[0] ?? 0)}`,
    ...breaks.slice(1).map((b, i) => `${write(breaks[i] ?? 0)} to ${write(b)}`),
    `${write(breaks[breaks.length - 1] ?? 0)} and over`,
  ]
}
