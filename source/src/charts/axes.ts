// Value scales with "nice" ticks: steps of 1, 2, 2.5 or 5 times a power of ten.

export interface LinearScale {
  min: number
  max: number
  step: number
  ticks: number[]
}

const NICE = [1, 2, 2.5, 5, 10]

/** Cleans floating-point noise such as 0.30000000000000004. */
export const tidy = (value: number): number => Number.parseFloat(value.toPrecision(12))

/** The smallest nice step at least as large as raw. */
export function niceStep(raw: number): number {
  if (!(raw > 0) || !Number.isFinite(raw)) return 1
  const power = 10 ** Math.floor(Math.log10(raw))
  for (const m of NICE) {
    const step = tidy(m * power)
    if (step >= raw * (1 - 1e-9)) return step
  }
  return tidy(10 * power)
}

/** Decimals a step needs: 2.5 needs one, 0.25 two, 50 none. */
export function stepDecimals(step: number): number {
  for (let d = 0; d < 10; d += 1) if (Math.abs(Number(step.toFixed(d)) - step) < 1e-9 * Math.max(1, step)) return d
  return 10
}

export interface ScaleOptions {
  /** The most ticks that fit; at least two. */
  maxTicks?: number
  /** Fixed ends, such as an axis range typed by the author. */
  min?: number
  max?: number
  /** Keep zero in the range (bars must start from zero). Default true. */
  includeZero?: boolean
}

/**
 * A scale that covers lo to hi with round ticks. Fixed ends are kept exactly; ticks are the multiples of the step
 * between them.
 */
export function niceScale(lo: number, hi: number, options: ScaleOptions = {}): LinearScale {
  const maxTicks = Math.max(2, Math.floor(options.maxTicks ?? 6))
  let a = Number.isFinite(lo) ? lo : 0
  let b = Number.isFinite(hi) ? hi : 1
  if (options.includeZero ?? true) {
    a = Math.min(0, a)
    b = Math.max(0, b)
  }
  if (options.min !== undefined) a = options.min
  if (options.max !== undefined) b = options.max
  if (a > b) [a, b] = [b, a]
  if (a === b) {
    if (a === 0) b = 1
    else if (a > 0) a = options.min ?? 0
    else b = options.max ?? 0
    if (a === b) b = a + Math.abs(a || 1)
  }
  let step = niceStep((b - a) / (maxTicks - 1))
  const lower = (s: number): number => options.min ?? tidy(Math.floor(a / s + 1e-9) * s)
  const upper = (s: number): number => options.max ?? tidy(Math.ceil(b / s - 1e-9) * s)
  // Rounding the ends out can add a tick beyond the limit; take the next step up when it does.
  while (Math.round((upper(step) - lower(step)) / step) + 1 > maxTicks) step = niceStep(step * 1.0001)
  const min = lower(step)
  const max = upper(step)
  const ticks: number[] = []
  for (let k = Math.ceil(min / step - 1e-9); k * step <= max + step * 1e-9; k += 1) ticks.push(tidy(k * step))
  return { min, max, step, ticks }
}

/** Zero stays on an axis unless the values sit far from it, where it would flatten the data. */
export function nearZero(lo: number, hi: number): boolean {
  if (lo >= 0) return lo <= hi * 0.5
  if (hi <= 0) return hi >= lo * 0.5
  return true
}

/** Maps a value on the scale to a coordinate between start (the minimum) and end (the maximum). */
export function project(scale: LinearScale, value: number, start: number, end: number): number {
  const span = scale.max - scale.min || 1
  return start + ((value - scale.min) / span) * (end - start)
}
