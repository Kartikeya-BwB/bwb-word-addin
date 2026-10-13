// Table data for the PowerPoint table tools, as plain arrays: the BwB look for each cell, totals, sorting, column
// widths and text pasted from Excel. Figures are read by the same rules as the Word table tools.

import { SLIDE_TYPE } from '../../brand/slideGrid.ts'
import { COLOURS } from '../../brand/tokens.ts'
import { parsePasted } from '../../charts/data.ts'
import { cmToPt, ptToCm } from '../../geometry/index.ts'
import { figureColumns, hasTotalRow } from '../../word/tables.ts'
import { sortRows, totalOf } from '../../word/tools/figures.ts'
import { UserFacingError } from '../../word/run.ts'

export { transpose } from '../../word/tools/figures.ts'

export interface CellLook {
  fill: string
  colour: string
  bold: boolean
  align: 'Left' | 'Right'
}

export interface TablePlan {
  /** Columns of figures, right-aligned. */
  figures: number[]
  /** The last row is a total row. */
  total: boolean
  cells: CellLook[][]
}

/** Table text: Aptos at the small slide size. */
export const TABLE_FONT = { name: SLIDE_TYPE.small.font, size: SLIDE_TYPE.small.size } as const
/** Light grey rules on every side of every cell. */
export const TABLE_RULE = { colour: COLOURS.lightGrey, weight: 0.75 } as const

/** The BwB look: a navy header row in white bold, figures right-aligned, a total row off-white and bold. */
export function planTable(values: readonly string[][]): TablePlan {
  const rows = values.map((r) => [...r])
  const figures = figureColumns(rows, 1)
  const total = hasTotalRow(rows, 1)
  const last = rows.length - 1
  const cells = rows.map((row, r) => row.map((_, c): CellLook => {
    const header = r === 0
    const isTotal = total && r === last
    return {
      fill: header ? COLOURS.navy : isTotal ? COLOURS.offWhite : COLOURS.white,
      colour: header ? COLOURS.white : COLOURS.black,
      bold: header || isTotal,
      align: figures.includes(c) ? 'Right' : 'Left',
    }
  }))
  return { figures, total, cells }
}

export interface TotalResult {
  values: string[][]
  /** False when the last row was already a total and was brought up to date. */
  added: boolean
  /** Columns that were added up. */
  summed: number[]
  /** Columns left blank, with the reason: "Year (years)". */
  skipped: string[]
}

/** Adds a total row that adds up each column of figures, or updates the totals in an existing one. */
export function withTotalRow(values: readonly string[][]): TotalResult {
  const rows = values.map((r) => [...r])
  const header = rows[0] ?? []
  const existing = hasTotalRow(rows, 1)
  const body = rows.slice(1, existing ? -1 : undefined)
  const width = Math.max(0, ...rows.map((r) => r.length))
  const total = existing ? [...(rows[rows.length - 1] ?? [])] : Array.from({ length: width }, (_, c) => (c === 0 ? 'Total' : ''))
  const summed: number[] = []
  const skipped: string[] = []
  for (let c = 1; c < width; c += 1) {
    const result = totalOf(body.map((r) => r[c] ?? ''), header[c] ?? '')
    if (result.ok) {
      total[c] = result.text
      summed.push(c)
    } else {
      total[c] = ''
      if (result.reason !== 'no figures') skipped.push(`${header[c] || `column ${c + 1}`} (${result.reason})`)
    }
  }
  const out = existing ? [...rows.slice(0, -1), total] : [...rows, total]
  return { values: out, added: !existing, summed, skipped }
}

/** Sorts the body by one column: by value for a column of figures, otherwise A to Z. The header and a total row stay. */
export function sortedValues(values: readonly string[][], column: number, descending: boolean): { values: string[][]; numeric: boolean } {
  const rows = values.map((r) => [...r])
  const numeric = figureColumns(rows, 1).includes(column)
  return { values: sortRows(rows, { column, descending, numeric, headerRows: 1, keepLast: hasTotalRow(rows, 1) }), numeric }
}

/** Equal shares of a width. */
export function equalWidths(total: number, columns: number): number[] {
  return Array.from({ length: columns }, () => total / Math.max(1, columns))
}

/** The widest a column can be: the width of a 16:9 slide. */
const MAX_CM = 33.87

/** Column widths typed in centimetres ("3, 4.5, 2" or "3 4.5 2"), one for each column, in points. */
export function parseWidths(text: string, columns: number): number[] {
  const parts = text.split(/[\s;,]+/).filter(Boolean)
  if (parts.length !== columns) {
    throw new UserFacingError(`Give ${columns} widths in centimetres, one for each column: you gave ${parts.length}.`)
  }
  return parts.map((p) => {
    const cm = Number(p)
    if (!Number.isFinite(cm) || cm <= 0 || cm > MAX_CM) throw new UserFacingError(`“${p}” is not a column width in centimetres.`)
    return cmToPt(cm)
  })
}

/** "3, 4.5 and 2 cm" for a message. */
export function describeWidths(widths: readonly number[]): string {
  const cms = widths.map((w) => String(Math.round(ptToCm(w) * 100) / 100))
  const list = cms.length > 1 ? `${cms.slice(0, -1).join(', ')} and ${cms[cms.length - 1] ?? ''}` : (cms[0] ?? '')
  return `${list} cm`
}

/** The most a slide table takes before it stops being readable. */
export const MAX_ROWS = 30
export const MAX_COLUMNS = 12

/** Cells copied from Excel (tab-separated) or a CSV file, as rows of text. */
export function pastedGrid(text: string): string[][] {
  const { grid } = parsePasted(text)
  const rows = grid.filter((r) => r.some((c) => c.trim()))
  const columns = Math.max(0, ...rows.map((r) => r.length))
  if (!rows.length || !columns) throw new UserFacingError('Paste cells copied from Excel first.')
  if (rows.length > MAX_ROWS || columns > MAX_COLUMNS) {
    throw new UserFacingError(`That is ${rows.length} rows by ${columns} columns, too big to read on a slide. Keep it to ${MAX_ROWS} rows and ${MAX_COLUMNS} columns.`)
  }
  return rows.map((r) => Array.from({ length: columns }, (_, c) => (r[c] ?? '').trim()))
}
