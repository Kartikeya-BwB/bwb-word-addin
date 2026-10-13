// Dates for Gantt charts. A date is a day number (days since 1 January 1970, UTC), so time zones never shift a
// task by a day. Text can be a full date (5 October 2026, 2026-10-05, 05/10/2026), a month (October 2026), a quarter
// (Q4 2026), a half (H2 2026) or a year; a month as an end date means the end of that month.

export type Precision = 'day' | 'month' | 'quarter' | 'half' | 'year'

export interface DateSpan {
  /** First day the text covers. */
  first: number
  /** Last day the text covers, inclusive. */
  last: number
  precision: Precision
}

export const MONTH_NAMES = [
  'January', 'February', 'March', 'April', 'May', 'June', 'July', 'August', 'September', 'October', 'November', 'December',
]

const DAY_MS = 86_400_000

export const dayNumber = (year: number, month: number, day: number): number => Math.round(Date.UTC(year, month, day) / DAY_MS)

/** Year, month (0 to 11) and day of a day number. */
export function civil(day: number): { year: number; month: number; day: number } {
  const date = new Date(day * DAY_MS)
  return { year: date.getUTCFullYear(), month: date.getUTCMonth(), day: date.getUTCDate() }
}

const daysIn = (year: number, month: number): number => dayNumber(year, month + 1, 1) - dayNumber(year, month, 1)

function monthIndex(word: string): number | undefined {
  const w = word.toLowerCase().replace(/\.$/, '')
  if (w.length < 3) return undefined
  const i = MONTH_NAMES.findIndex((m) => m.toLowerCase().startsWith(w))
  return i >= 0 && (w.length === 3 || w === 'sept' || MONTH_NAMES[i]?.toLowerCase() === w) ? i : undefined
}

function fullYear(text: string): number {
  const y = Number(text)
  return text.length <= 2 ? (y < 70 ? 2000 + y : 1900 + y) : y
}

function day(year: number, month: number, date: number): DateSpan | null {
  if (month < 0 || month > 11 || date < 1 || date > daysIn(year, month)) return null
  const n = dayNumber(year, month, date)
  return { first: n, last: n, precision: 'day' }
}

function months(year: number, month: number, count: number, precision: Precision): DateSpan {
  return { first: dayNumber(year, month, 1), last: dayNumber(year, month + count, 1) - 1, precision }
}

/** Reads a date or period, or returns null. */
export function parseDate(input: string): DateSpan | null {
  const s = input.trim().replace(/\s+/g, ' ')
  if (!s) return null
  let m = /^(\d{4})[-/.](\d{1,2})[-/.](\d{1,2})(?:[T ][\d:.]+Z?)?$/.exec(s)
  if (m) return day(Number(m[1]), Number(m[2]) - 1, Number(m[3]))
  m = /^(\d{4})-(\d{1,2})$/.exec(s)
  if (m && Number(m[2]) >= 1 && Number(m[2]) <= 12) return months(Number(m[1]), Number(m[2]) - 1, 1, 'month')
  m = /^(\d{1,2})(?:st|nd|rd|th)?[ \-./]+([A-Za-z]+)\.?,?[ \-./]+(\d{4}|\d{2})$/.exec(s)
  if (m) {
    const month = monthIndex(m[2] ?? '')
    return month === undefined ? null : day(fullYear(m[3] ?? ''), month, Number(m[1]))
  }
  m = /^([A-Za-z]+)\.? (\d{1,2})(?:st|nd|rd|th)?,? (\d{4})$/.exec(s)
  if (m) {
    const month = monthIndex(m[1] ?? '')
    return month === undefined ? null : day(Number(m[3]), month, Number(m[2]))
  }
  m = /^([A-Za-z]+)\.?[ \-/']+(\d{4}|\d{2})$/.exec(s)
  if (m) {
    const month = monthIndex(m[1] ?? '')
    if (month !== undefined) return months(fullYear(m[2] ?? ''), month, 1, 'month')
  }
  const q1 = /^Q([1-4])[ \-/']*(?:FY)?(\d{4}|\d{2})$/i.exec(s)
  const q2 = /^(?:FY)?(\d{4}) ?-? ?Q([1-4])$/i.exec(s)
  if (q1 || q2) {
    const quarter = Number(q1 ? q1[1] : q2?.[2])
    return months(fullYear((q1 ? q1[2] : q2?.[1]) ?? ''), (quarter - 1) * 3, 3, 'quarter')
  }
  m = /^H([12])[ \-/']*(\d{4}|\d{2})$/i.exec(s)
  if (m) return months(fullYear(m[2] ?? ''), (Number(m[1]) - 1) * 6, 6, 'half')
  m = /^(\d{4})$/.exec(s)
  if (m && Number(m[1]) >= 1900 && Number(m[1]) <= 2200) return months(Number(m[1]), 0, 12, 'year')
  m = /^(\d{1,2})[/.\-](\d{1,2})[/.\-](\d{4}|\d{2})$/.exec(s)
  if (m) {
    const [a, b] = [Number(m[1]), Number(m[2])]
    // Day first, as in the UK; a first number above 12 can only be a day, a second above 12 only a US-style day.
    return b > 12 && a <= 12 ? day(fullYear(m[3] ?? ''), a - 1, b) : day(fullYear(m[3] ?? ''), b - 1, a)
  }
  m = /^(\d{5})(?:\.\d+)?$/.exec(s)
  if (m) {
    // An Excel serial number (the 1900 date system), as cells formatted as numbers copy.
    const n = dayNumber(1899, 11, 30) + Number(m[1])
    return { first: n, last: n, precision: 'day' }
  }
  return null
}

/** Month names cut to three letters ("Sep 2026"), when every label is a month name; otherwise null. */
export function abbreviateMonths(labels: string[]): string[] | null {
  const out = labels.map((label) => {
    const m = /^([A-Za-z]+)(\s+'?\d{2,4})?$/.exec(label.trim())
    const name = MONTH_NAMES.find((n) => n.toLowerCase() === (m?.[1] ?? '').toLowerCase())
    return m && name ? `${name.slice(0, 3)}${m[2] ?? ''}` : null
  })
  return out.every((l): l is string => l !== null) && out.some((l, i) => l !== labels[i]) ? out : null
}

/** "5 Oct", "5 October 2026", or "October 2026". */
export function formatDay(n: number, style: 'short' | 'long' | 'month' = 'long'): string {
  const { year, month, day: d } = civil(n)
  const name = MONTH_NAMES[month] ?? ''
  if (style === 'short') return `${d} ${name.slice(0, 3)}`
  if (style === 'month') return `${name} ${year}`
  return `${d} ${name} ${year}`
}

/** Today in Europe/London as a day number. */
export function londonToday(now: Date = new Date()): number {
  const parts = new Intl.DateTimeFormat('en-GB', { timeZone: 'Europe/London', year: 'numeric', month: 'numeric', day: 'numeric' })
    .formatToParts(now)
  const part = (type: string): number => Number(parts.find((p) => p.type === type)?.value ?? 0)
  return dayNumber(part('year'), part('month') - 1, part('day'))
}

/** ISO text of a day number, "2026-10-13". */
export function isoDay(n: number): string {
  const { year, month, day: d } = civil(n)
  return `${year}-${String(month + 1).padStart(2, '0')}-${String(d).padStart(2, '0')}`
}

export type TimeUnit = 'month' | 'quarter' | 'year'

export interface Period {
  /** First day. */
  start: number
  /** First day after the period. */
  end: number
  year: number
  /** Labels from longest to shortest: "Jan", "J"; "Q1"; "2027". */
  labels: string[]
}

const unitMonths: Record<TimeUnit, number> = { month: 1, quarter: 3, year: 12 }

/** First day of the period holding a day. */
export function periodStart(n: number, unit: TimeUnit): number {
  const { year, month } = civil(n)
  const size = unitMonths[unit]
  return dayNumber(year, Math.floor(month / size) * size, 1)
}

/** Periods from the one holding `from` to the one holding `to`, inclusive. */
export function periods(from: number, to: number, unit: TimeUnit): Period[] {
  const out: Period[] = []
  const size = unitMonths[unit]
  let start = periodStart(from, unit)
  while (start <= to && out.length < 600) {
    const { year, month } = civil(start)
    const end = dayNumber(year, month + size, 1)
    const name = MONTH_NAMES[month] ?? ''
    const labels =
      unit === 'month' ? [name.slice(0, 3), name.slice(0, 1)] : unit === 'quarter' ? [`Q${Math.floor(month / 3) + 1}`] : [`${year}`]
    out.push({ start, end, year, labels })
    start = end
  }
  return out
}

/** Months for plans up to 15 months, quarters up to four years, then years. */
export function chooseUnit(from: number, to: number): TimeUnit {
  const span = (to - from) / 30.44
  return span <= 15 ? 'month' : span <= 48 ? 'quarter' : 'year'
}
