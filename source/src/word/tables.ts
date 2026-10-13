// Formatting an existing table, such as one pasted from Excel or another document, as a BwB table.

import { TABLE_STYLES } from '../brand/styles.ts'
import { UserFacingError, runWord } from './run.ts'
import { withBwbStyle } from './styles.ts'

export type TableKind = keyof typeof TABLE_STYLES

/** Figures as BwB tables show them: 1,234 · 12.5% · (300) · -4 · USD 250m · 3.2bn · 1.5x · 2030. */
const FIGURE = /^[-+\u2212(]?\s*(?:[A-Z]{3}\s?)?[-+\u2212]?\d[\d,.\u00A0\u202F ]*(?:%|pp|k|m|bn|x)?\)?\*?$/i
/** Marks for a missing value, which do not stop a column counting as figures. */
const EMPTY_MARK = /^(?:|-|\u2013|\u2014|n\/?a|\.\.|\u2026)$/i

export function isFigure(text: string): boolean {
  return FIGURE.test(text.trim())
}

/** Columns whose body cells are all figures (missing-value marks aside), which BwB tables right-align. */
export function figureColumns(values: string[][], headerRows: number): number[] {
  const body = values.slice(headerRows)
  const width = Math.max(0, ...values.map((r) => r.length))
  const out: number[] = []
  for (let c = 0; c < width; c++) {
    const cells = body.map((r) => (r[c] ?? '').trim()).filter((t) => !EMPTY_MARK.test(t))
    if (cells.length && cells.every(isFigure)) out.push(c)
  }
  return out
}

/** True when the last row is a total row ("Total", "Total funding", "Grand total"). */
export function hasTotalRow(values: string[][], headerRows: number): boolean {
  const last = values[values.length - 1]
  return values.length > headerRows + 1 && !!last && /^\s*(?:grand\s+|sub-?)?total\b/i.test(last[0] ?? '')
}

/**
 * Gives the table at the cursor a BwB table style, repeats its header row on each page, right-aligns columns
 * of figures and styles a total row. Returns a sentence saying what changed.
 */
export async function formatTableAtCursor(kind: TableKind = 'standard'): Promise<string> {
  const style = TABLE_STYLES[kind].name
  const { result, importedStyles } = await withBwbStyle(style, () =>
    runWord(async (context) => {
      const table = context.document.getSelection().parentTableOrNullObject
      table.load('isNullObject,values,rowCount,headerRowCount')
      await context.sync()
      if (table.isNullObject) throw new UserFacingError('Put the cursor in a table first.')

      const headerRows = Math.max(1, table.headerRowCount)
      const values = table.values
      const total = kind === 'standard' && hasTotalRow(values, headerRows)
      table.style = style
      table.headerRowCount = headerRows
      table.styleFirstColumn = kind === 'comparison'
      table.styleTotalRow = total
      table.styleBandedRows = false
      table.styleBandedColumns = false

      const columns = figureColumns(values, headerRows)
      const cells = columns.flatMap((c) =>
        values.map((row, r) => (c < row.length ? table.getCellOrNullObject(r, c) : null)).filter((x) => x !== null),
      )
      cells.forEach((cell) => cell.load('isNullObject'))
      await context.sync()
      for (const cell of cells) if (!cell.isNullObject) cell.horizontalAlignment = 'Right'
      await context.sync()
      return { columns: columns.length, total }
    }),
  )
  const parts = [`${TABLE_STYLES[kind].name} applied`]
  if (result.columns) parts.push(`${result.columns} column${result.columns === 1 ? '' : 's'} of figures right-aligned`)
  if (result.total) parts.push('total row styled')
  return `${parts.join(', ')}.${importedStyles ? ' BwB styles were added to this document.' : ''}`
}
