// Numbers the BwB way: thousands separators, a currency as an ISO code before the number with a space (USD 250m),
// a scale written straight after the number (250m, 1.2bn), the % sign in charts, and a true minus sign.

export type Scale = '' | 'k' | 'm' | 'bn' | 'tn'

export const SCALES: readonly Scale[] = ['', 'k', 'm', 'bn', 'tn']

export const SCALE_FACTOR: Record<Scale, number> = { '': 1, k: 1e3, m: 1e6, bn: 1e9, tn: 1e12 }

/**
 * How a chart writes its numbers. Values are already in the units shown: with scale "m", 250 is written 250m, and
 * with percent, 12 is written 12%.
 */
export interface NumberFormat {
  /** Fixed decimal places; automatic when absent. */
  decimals?: number
  scale?: Scale
  percent?: boolean
  /** ISO 4217 code such as USD; an empty string means none. */
  currency?: string
}

export const MINUS = '\u{2212}'

/** "1234567" as "1,234,567". */
export function groupThousands(digits: string): string {
  return digits.replace(/\B(?=(\d{3})+(?!\d))/g, ',')
}

/** The fewest decimals (up to max) that write the value without losing anything. */
export function decimalsNeeded(value: number, max = 4): number {
  for (let d = 0; d < max; d += 1) {
    if (Math.abs(Number(value.toFixed(d)) - value) <= 1e-9 * Math.max(1, Math.abs(value))) return d
  }
  return max
}

/**
 * Decimals for a set of values written together, such as the data labels of one chart: enough for the most precise
 * value, but at most two, one from 10 and none from 1,000, so labels stay short and line up.
 */
export function autoDecimals(values: number[]): number {
  const finite = values.filter(Number.isFinite)
  if (!finite.length) return 0
  const largest = Math.max(...finite.map(Math.abs))
  const cap = largest >= 1000 ? 0 : largest >= 10 ? 1 : 2
  return Math.min(cap, Math.max(0, ...finite.map((v) => decimalsNeeded(v, 2))))
}

export interface FormatOptions {
  /** Overrides the format's decimals. */
  decimals?: number
  /** Writes a plus sign before positive values (waterfall changes). */
  plus?: boolean
  /** Leaves out the currency and scale, for labels whose units are written once elsewhere ("USD m"). */
  bare?: boolean
}

function suffix(format: NumberFormat): string {
  return format.percent ? '%' : (format.scale ?? '')
}

/** One number written the BwB way: "USD 1,250m", "12.5%", or a negative with a true minus sign. */
export function formatNumber(value: number, format: NumberFormat = {}, options: FormatOptions = {}): string {
  if (!Number.isFinite(value)) return ''
  const decimals = Math.min(6, Math.max(0, Math.round(options.decimals ?? format.decimals ?? autoDecimals([value]))))
  const fixed = Math.abs(value).toFixed(decimals)
  const [whole = '0', fraction] = fixed.split('.')
  const isZero = Number(fixed) === 0
  const sign = isZero ? '' : value < 0 ? MINUS : options.plus ? '+' : ''
  const unit = options.bare ? (format.percent ? '%' : '') : suffix(format)
  const body = `${sign}${groupThousands(whole)}${fraction ? `.${fraction}` : ''}${unit}`
  const currency = (format.currency ?? '').trim().toUpperCase()
  return currency && !options.bare ? `${currency} ${body}` : body
}

/** A short description of the units, for alt text and hints: "USD m", "%", "bn". */
export function unitsLabel(format: NumberFormat): string {
  const currency = (format.currency ?? '').trim().toUpperCase()
  const scale = format.percent ? '%' : (format.scale ?? '')
  return [currency, scale].filter(Boolean).join(' ')
}
