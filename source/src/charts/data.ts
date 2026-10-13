// Chart data pasted from Excel: tab-separated (what Excel puts on the clipboard), with comma or semicolon separated
// text as a fallback. The first row holds the series names and the first column the categories.

import type { NumberFormat, Scale } from './format.ts'
import { SCALE_FACTOR } from './format.ts'
import type { Grid } from './grid.ts'
import { normaliseGrid, trimGrid } from './grid.ts'

export type { Grid } from './grid.ts'
export {
  addColumn, addRow, cellAt, gridToText, moveColumn, normaliseGrid, pasteBlock, removeColumn, removeRow, setCell,
  transpose, trimGrid,
} from './grid.ts'

export interface ParsedNumber {
  /** The number in its own units (1.2 for "USD 1.2bn"), or null for a gap. */
  value: number | null
  /** False when the cell holds text that is not a number; it is then treated as a gap. */
  ok: boolean
  currency?: string
  scale?: Scale
  percent?: boolean
}

const GAP = /^(?:n\/?a|n\.a\.|na|-+|\u{2013}|\u{2014}|\.\.|\u{2026}|#n\/a|#div\/0!|#value!|#ref!|#num!|null|none)$/iu
const SYMBOLS: Record<string, string> = { $: 'USD', US$: 'USD', '\u{00A3}': 'GBP', '\u{20AC}': 'EUR' }
const SCALE_WORDS: Array<[RegExp, Scale]> = [
  [/^(?:k|thousand|thousands)$/i, 'k'],
  [/^(?:m|mn|mm|mln|million|millions)$/i, 'm'],
  [/^(?:bn|b|billion|billions)$/i, 'bn'],
  [/^(?:tn|t|trillion|trillions)$/i, 'tn'],
]

function scaleOf(word: string): Scale | undefined {
  return SCALE_WORDS.find(([pattern]) => pattern.test(word))?.[1]
}

/** Digits with optional thousands separators and decimals, in either convention. */
function parseDigits(body: string, decimalComma: boolean): number | null {
  const s = body.replace(/[\s\u{00A0}\u{202F}\u{2009}]/gu, '')
  const point = /^(?:\d{1,3}(?:,\d{3})+|\d+)?(?:\.\d+)?$/
  const comma = /^(?:\d{1,3}(?:\.\d{3})+|\d+)?(?:,\d+)?$/
  if (!/\d/.test(s)) return null
  if (decimalComma ? comma.test(s) : point.test(s)) {
    return Number(decimalComma ? s.replace(/\./g, '').replace(',', '.') : s.replace(/,/g, ''))
  }
  // "1,5" in a point-decimal table is a decimal comma; "1.234,5" is the other convention whatever the table says.
  if (!decimalComma && comma.test(s) && !/^\d{1,3},\d{3}$/.test(s)) return Number(s.replace(/\./g, '').replace(',', '.'))
  return null
}

/**
 * Reads a number the way people type it in Excel and reports: "1,234.5", "(300)", "-4", "12%", "USD 1.2bn", "$5m",
 * "3.5k". Empty cells, "n/a" and dashes are gaps.
 */
export function parseNumber(input: string, options: { decimalComma?: boolean } = {}): ParsedNumber {
  let s = input.replace(/[\u{00A0}\u{202F}]/gu, ' ').trim()
  if (!s || GAP.test(s)) return { value: null, ok: true }
  let negative = false
  let currency: string | undefined
  let scale: Scale | undefined
  let percent = false
  for (let guard = 0; guard < 6; guard += 1) {
    const before = s
    const paren = /^\((.*)\)$/.exec(s)
    if (paren) {
      negative = !negative
      s = (paren[1] ?? '').trim()
    }
    const sign = /^([+\-\u{2212}\u{2013}])\s*/u.exec(s)
    if (sign) {
      if (sign[1] !== '+') negative = !negative
      s = s.slice(sign[0].length)
    }
    const prefix = /^(US\$|[A-Z]{3}(?![a-z])|[$\u{00A3}\u{20AC}])\s*/u.exec(s)
    if (prefix && !currency) {
      const code = prefix[1] ?? ''
      currency = SYMBOLS[code] ?? code
      s = s.slice(prefix[0].length)
    }
    const after = /\s*(?:\b([A-Z]{3})|([$\u{00A3}\u{20AC}]))$/u.exec(s)
    if (after && !currency && /\d/.test(s.slice(0, after.index))) {
      currency = after[1] ?? SYMBOLS[after[2] ?? ''] ?? ''
      s = s.slice(0, after.index)
    }
    if (/\s*%$/.test(s)) {
      percent = true
      s = s.replace(/\s*%$/, '')
    }
    const word = /(\d)\s*([A-Za-z]+)$/.exec(s)
    const found = word ? scaleOf(word[2] ?? '') : undefined
    if (word && found && !scale) {
      scale = found
      s = s.slice(0, s.length - (word[0].length - 1))
    }
    s = s.trim()
    if (s === before) break
  }
  const digits = parseDigits(s, options.decimalComma ?? false)
  if (digits === null || !Number.isFinite(digits)) return { value: null, ok: false }
  return {
    value: negative ? -digits : digits,
    ok: true,
    ...(currency ? { currency } : {}),
    ...(scale ? { scale } : {}),
    ...(percent ? { percent } : {}),
  }
}

export type Delimiter = '\t' | ';' | ','

/** Splits delimited text, honouring double quotes the way Excel and CSV files use them. */
function splitRows(text: string, delimiter: string): Grid {
  const rows: Grid = []
  let row: string[] = []
  let cell = ''
  let quoted = false
  for (let i = 0; i < text.length; i += 1) {
    const ch = text[i] ?? ''
    if (quoted) {
      if (ch === '"' && text[i + 1] === '"') {
        cell += '"'
        i += 1
      } else if (ch === '"') quoted = false
      else cell += ch
    } else if (ch === '"' && cell.trim() === '') {
      quoted = true
      cell = ''
    } else if (ch === delimiter) {
      row.push(cell)
      cell = ''
    } else if (ch === '\n') {
      row.push(cell)
      rows.push(row)
      row = []
      cell = ''
    } else cell += ch
  }
  row.push(cell)
  rows.push(row)
  return rows
}

function chooseDelimiter(text: string): Delimiter | null {
  if (text.includes('\t')) return '\t'
  const lines = text.split('\n').filter((l) => l.trim())
  for (const d of [';', ','] as const) {
    const counts = lines.map((l) => splitRows(l, d)[0]?.length ?? 1)
    if (counts.length && counts.every((c) => c > 1 && c === counts[0])) return d
  }
  for (const d of [';', ','] as const) if (lines.some((l) => l.includes(d))) return d
  return null
}

export interface PastedTable {
  grid: Grid
  delimiter: Delimiter | null
  /** True for European number style (1.234,5), which comes with semicolon-separated text. */
  decimalComma: boolean
}

/** Parses text pasted from Excel or a CSV file into a rectangular grid of cell text. */
export function parsePasted(text: string): PastedTable {
  const clean = text.replace(/\r\n?/g, '\n').replace(/\n+$/, '')
  const delimiter = chooseDelimiter(clean)
  const rows = delimiter ? splitRows(clean, delimiter) : clean.split('\n').map((l) => [l])
  const grid = trimGrid(normaliseGrid(rows.map((r) => r.map((c) => c.trim()))))
  const decimalComma = delimiter === ';' && grid.some((r) => r.some((c) => /\d,\d/.test(c) && !/\d\.\d+$/.test(c)))
  return { grid, delimiter, decimalComma }
}

export interface Series {
  name: string
  values: Array<number | null>
}

export interface SeriesTable {
  /** Heading of the first column, such as "Year". */
  corner: string
  categories: string[]
  series: Series[]
  /** Units read from the cells: currency, scale and per cent. */
  hint: NumberFormat
  /** Cells that held text other than a number; they are drawn as gaps. */
  invalid: Array<{ row: number; column: number; text: string }>
}

function mostCommon<T>(items: T[]): T | undefined {
  const counts = new Map<T, number>()
  for (const item of items) counts.set(item, (counts.get(item) ?? 0) + 1)
  let best: T | undefined
  let bestCount = 0
  for (const [item, count] of counts) if (count > bestCount) [best, bestCount] = [item, count]
  return best
}

/** Units the cells agree on, for the number format: "USD 1.2bn" and "USD 800m" give USD and bn. */
export function unitsHint(cells: ParsedNumber[]): NumberFormat {
  const numbers = cells.filter((c) => c.value !== null)
  const currencies = numbers.map((c) => c.currency).filter((c): c is string => Boolean(c))
  const scales = numbers.map((c) => c.scale).filter((c): c is Scale => Boolean(c))
  const scale = mostCommon(scales)
  const percent = numbers.length > 0 && numbers.filter((c) => c.percent).length * 2 > numbers.length
  return {
    ...(currencies.length && new Set(currencies).size === 1 ? { currency: currencies[0] ?? '' } : {}),
    ...(scale && !percent ? { scale } : {}),
    ...(percent ? { percent } : {}),
  }
}

/**
 * Reads a grid as categories (first column) against series (first row). Numbers are put in the units the column
 * mostly uses, so "800m" next to "1.2bn" is 0.8 when the chart is in billions.
 */
export function toSeriesTable(input: Grid, options: { decimalComma?: boolean } = {}): SeriesTable {
  const grid = normaliseGrid(input)
  const header = grid[0] ?? []
  const body = grid.slice(1).filter((row) => row.some((c) => c.trim()))
  const parsed = body.map((row) => row.slice(1).map((c) => parseNumber(c, options)))
  const hint = unitsHint(parsed.flat())
  const factor = SCALE_FACTOR[hint.scale ?? '']
  const invalid: SeriesTable['invalid'] = []
  const series: Series[] = header.slice(1).map((name, s) => ({
    name: name.trim() || `Series ${s + 1}`,
    values: parsed.map((row, r) => {
      const cell = row[s]
      if (!cell) return null
      if (!cell.ok) invalid.push({ row: r + 1, column: s + 1, text: body[r]?.[s + 1] ?? '' })
      if (cell.value === null) return null
      const own = cell.scale ? SCALE_FACTOR[cell.scale] : factor
      return (cell.value * own) / factor
    }),
  }))
  return {
    corner: (header[0] ?? '').trim(),
    categories: body.map((row, r) => (row[0] ?? '').trim() || `${r + 1}`),
    series,
    hint,
    invalid,
  }
}
