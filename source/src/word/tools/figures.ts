// Figures as BwB tables write them (1,234; 12.5%; (300); USD 250m; 1.2bn), read and written for totals and
// sorting; and table data as plain arrays: sorting rows, transposing, and splitting delimited lines into cells.

import { isFigure } from '../tables.ts'

type Scale = '' | 'k' | 'm' | 'bn'

const SCALES: Record<Scale, number> = { '': 1, k: 1e3, m: 1e6, bn: 1e9 }
const ORDER: Scale[] = ['', 'k', 'm', 'bn']

export interface Figure {
  /** The value in units: USD 250m is 250 000 000. */
  value: number
  /** Decimal places as written. */
  decimals: number
  scale: Scale
  suffix: '' | '%' | 'pp' | 'x'
  /** ISO currency code, upper case, or ''. */
  currency: string
  /** A negative written as (300). */
  parentheses: boolean
  /** The minus sign written, if any. */
  minus: '' | '-' | '\u2212'
}

/** A missing value, which a column of figures may hold: -, n/a, .., an en or em dash, an ellipsis. */
export const isMissing = (text: string): boolean => /^(?:|-|\u2013|\u2014|n\/?a|\.\.|\u2026)$/i.test(text.trim())

const PARTS = /^(\()?\s*([-+\u2212])?\s*(?:([A-Za-z]{3})\s?)?([-+\u2212])?(\d[\d,.\u00a0\u202f ]*?)\s*(%|pp|bn|k|m|x)?\s*(\))?$/i

/** Reads a figure, or null when the text is not one (or its digits are unclear, such as 1,5). */
export function parseFigure(text: string): Figure | null {
  const t = text.trim().replace(/\*+$/, '')
  if (!t || !isFigure(t)) return null
  const m = PARTS.exec(t)
  if (!m) return null
  const [, open, sign1, code, sign2, digitsRaw = '', suffixRaw = '', close] = m
  if (Boolean(open) !== Boolean(close)) return null
  const digits = digitsRaw.replace(/[\u00a0\u202f ]/g, '')
  if (!/^\d{1,3}(?:,\d{3})*(?:\.\d+)?$|^\d+(?:\.\d+)?$/.test(digits)) return null
  const number = Number(digits.replace(/,/g, ''))
  const sign = sign1 ?? sign2 ?? ''
  const negative = Boolean(open) || sign === '-' || sign === '\u2212'
  const unit = suffixRaw.toLowerCase()
  const scale: Scale = unit === 'k' || unit === 'm' || unit === 'bn' ? unit : ''
  const suffix = unit === '%' || unit === 'pp' || unit === 'x' ? unit : ''
  return {
    value: (negative ? -number : number) * SCALES[scale],
    decimals: digits.includes('.') ? (digits.split('.')[1]?.length ?? 0) : 0,
    scale,
    suffix,
    currency: code ? code.toUpperCase() : '',
    parentheses: Boolean(open),
    minus: sign === '-' || sign === '\u2212' ? sign : '',
  }
}

/** Decimal places a number needs, up to six. */
function decimalsOf(x: number): number {
  const s = Math.abs(x).toFixed(6).replace(/0+$/, '')
  const dot = s.indexOf('.')
  return dot < 0 || dot === s.length - 1 ? 0 : s.length - dot - 1
}

const withCommas = (digits: string): string => {
  const [whole = '', fraction] = digits.split('.')
  const grouped = whole.replace(/\B(?=(\d{3})+(?!\d))/g, ',')
  return fraction ? `${grouped}.${fraction}` : grouped
}

export type ColumnTotal = { ok: true; text: string; value: number } | { ok: false; reason: string }

/** Header words for columns that never add up, such as years, rates and rankings. */
const NOT_ADDITIVE = /\b(?:years?|dates?|rates?|ratios?|growth|change|yields?|rank(?:ing)?|average|mean|median|scores?|index)\b|^\s*(?:no\.?|#|ref\.?)\s*$/i

/**
 * The total of a column of figures, written the way the column writes them: USD 250m and USD 1.2bn give USD 1.45bn,
 * (300) and 500 give 200, shares adding to 100 per cent give 100%. Columns that cannot be added say why.
 */
export function totalOf(cells: string[], header = ''): ColumnTotal {
  const present = cells.map((c) => c.trim()).filter((c) => !isMissing(c))
  if (!present.length) return { ok: false, reason: 'no figures' }
  const figures = present.map(parseFigure)
  if (figures.some((f) => f === null)) return { ok: false, reason: 'text among the figures' }
  const list = figures.filter((f): f is Figure => f !== null)
  if (list.some((f) => f.suffix === 'x' || f.suffix === 'pp')) return { ok: false, reason: 'multiples or percentage points' }
  const percent = list.filter((f) => f.suffix === '%').length
  if (percent && percent !== list.length) return { ok: false, reason: 'mixed units' }
  const currencies = [...new Set(list.map((f) => f.currency).filter(Boolean))]
  if (currencies.length > 1) return { ok: false, reason: 'mixed currencies' }
  const nonZero = list.filter((f) => f.value !== 0)
  if (nonZero.some((f) => f.scale === '') && nonZero.some((f) => f.scale !== '')) return { ok: false, reason: 'mixed units' }
  const years = list.every((f) => !f.currency && !f.scale && !f.suffix && f.decimals === 0 && f.value >= 1900 && f.value <= 2100)
  if (years && present.every((c) => /^\d{4}$/.test(c))) return { ok: false, reason: 'years' }
  if (NOT_ADDITIVE.test(header)) return { ok: false, reason: 'does not add up' }

  const sum = list.reduce((a, f) => a + f.value, 0)
  if (percent && Math.abs(sum - 100) > 1) return { ok: false, reason: 'percentages that are not shares of a whole' }
  const scale = list.reduce<Scale>((s, f) => (ORDER.indexOf(f.scale) > ORDER.indexOf(s) ? f.scale : s), '')
  const decimals = Math.min(3, Math.max(...list.map((f) => (f.scale === scale ? f.decimals : decimalsOf(f.value / SCALES[scale])))))
  const amount = withCommas(Math.abs(sum / SCALES[scale]).toFixed(decimals)) + scale + (percent ? '%' : '')
  const currency = currencies[0] ? `${currencies[0]} ` : ''
  const rounded = Number((sum / SCALES[scale]).toFixed(decimals))
  let text = `${currency}${amount}`
  if (rounded < 0) {
    const minus = list.find((f) => f.minus)?.minus ?? '-'
    text = list.some((f) => f.parentheses) ? `(${text})` : `${minus}${text}`
  }
  return { ok: true, text, value: sum }
}

export interface SortOptions {
  column: number
  descending: boolean
  /** Sort by value (figures) rather than alphabetically. */
  numeric: boolean
  headerRows: number
  /** Keep the last row (a total) where it is. */
  keepLast: boolean
}

/** Sorts the body rows of a table by one column. Header rows and a total row stay put; empty cells go last. */
export function sortRows(values: string[][], o: SortOptions): string[][] {
  const head = values.slice(0, o.headerRows)
  const end = o.keepLast && values.length > o.headerRows + 1 ? values.length - 1 : values.length
  const body = values.slice(o.headerRows, end)
  const tail = values.slice(end)
  const key = (row: string[]): number | string | null => {
    const cell = (row[o.column] ?? '').trim()
    if (o.numeric) return parseFigure(cell)?.value ?? null
    return cell && !isMissing(cell) ? cell : null
  }
  const keyed = body.map((row, i) => ({ row, i, k: key(row) }))
  keyed.sort((a, b) => {
    if (a.k === null || b.k === null) return a.k === b.k ? a.i - b.i : a.k === null ? 1 : -1
    const c = typeof a.k === 'number' && typeof b.k === 'number'
      ? a.k - b.k
      : String(a.k).localeCompare(String(b.k), 'en-GB', { numeric: true, sensitivity: 'base' })
    return (o.descending ? -c : c) || a.i - b.i
  })
  return [...head, ...keyed.map((k) => k.row), ...tail]
}

/** Rows become columns. Short rows are padded with empty cells. */
export function transpose(values: string[][]): string[][] {
  const width = Math.max(0, ...values.map((r) => r.length))
  return Array.from({ length: width }, (_, c) => values.map((r) => r[c] ?? ''))
}

export type Delimiter = 'tab' | 'semicolon' | 'comma'

/** Splits on commas, keeping "quoted, text" together and leaving thousands separators (1,200) alone. */
function splitCommas(line: string): string[] {
  const thousands = (i: number): boolean => /\d/.test(line.charAt(i - 1)) && /^\d{3}(?!\d)/.test(line.slice(i + 1))
  const out: string[] = []
  let cell = ''
  let quoted = false
  for (let i = 0; i < line.length; i += 1) {
    const ch = line.charAt(i)
    if (ch === '"' && quoted && line.charAt(i + 1) === '"') {
      cell += '"'
      i += 1
    } else if (ch === '"') {
      quoted = !quoted
    } else if (ch === ',' && !quoted && !thousands(i)) {
      out.push(cell)
      cell = ''
    } else cell += ch
  }
  out.push(cell)
  return out
}

/**
 * Lines of text as table rows: split on tabs if any line has one, else on semicolons, else on commas. Rows are
 * padded to the same width and empty trailing columns dropped. Null when there is nothing to split on.
 */
export function rowsFromLines(lines: string[]): { delimiter: Delimiter; rows: string[][] } | null {
  const filled = lines.filter((l) => l.trim() !== '')
  const delimiter: Delimiter | null = filled.some((l) => l.includes('\t'))
    ? 'tab'
    : filled.some((l) => l.includes(';')) ? 'semicolon' : filled.some((l) => l.includes(',')) ? 'comma' : null
  if (!delimiter) return null
  const split = (l: string): string[] => (delimiter === 'tab' ? l.split('\t') : delimiter === 'semicolon' ? l.split(';') : splitCommas(l))
  const raw = filled.map((l) => split(l).map((c) => c.trim()))
  let width = Math.max(...raw.map((r) => r.length))
  while (width > 1 && raw.every((r) => !(r[width - 1] ?? ''))) width -= 1
  if (width < 2) return null
  return { delimiter, rows: raw.map((r) => Array.from({ length: width }, (_, c) => r[c] ?? '')) }
}
