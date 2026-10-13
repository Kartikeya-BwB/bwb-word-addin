// Table layout and totals for the table at the cursor: even columns, the full text width, centred, widths in
// centimetres, a total row that adds up the columns of figures, and deleting empty rows.

import { DATA_TABLE_STYLE_NAMES } from '../../brand/styles.ts'
import { UserFacingError, runWord } from '../run.ts'
import { figureColumns, hasTotalRow } from '../tables.ts'
import { cmToPt, count, formatCm, listOf } from './common.ts'
import { totalOf } from './figures.ts'
import { supportsDesktop, unavailable } from './support.ts'

/** The table holding the cursor, with the given properties loaded. */
export async function tableAtCursor(context: Word.RequestContext, props = 'values,headerRowCount,isUniform,style'): Promise<Word.Table> {
  const table = context.document.getSelection().parentTableOrNullObject
  table.load(`isNullObject,${props}`)
  await context.sync()
  if (table.isNullObject) throw new UserFacingError('Put the cursor in a table first.')
  return table
}

export const MERGED = 'This table has merged cells, so Word cannot do that here. Split the merged cells first (Layout, Split Cells).'

export async function distributeColumns(): Promise<string> {
  return runWord(async (context) => {
    const table = await tableAtCursor(context, 'isUniform')
    if (!table.isUniform) throw new UserFacingError(MERGED)
    table.distributeColumns()
    await context.sync()
    return 'Columns made the same width.'
  })
}

export async function fitTableToWindow(): Promise<string> {
  return runWord(async (context) => {
    const table = await tableAtCursor(context, 'rowCount')
    table.autoFitWindow()
    await context.sync()
    return 'Table stretched to the full text width.'
  })
}

/** Fits the columns to their contents. Needs Word for Windows or Mac (WordApiDesktop 1.4). */
export async function fitTableToContents(): Promise<string> {
  const reason = unavailable('fitToContents')
  if (reason || !supportsDesktop('1.4')) throw new UserFacingError(reason ?? 'Needs Word for Windows or Mac.')
  return runWord(async (context) => {
    const table = await tableAtCursor(context, 'rowCount')
    table.autoFitBehavior('Content')
    await context.sync()
    return 'Columns fitted to their contents.'
  })
}

export async function centreTable(): Promise<string> {
  return runWord(async (context) => {
    const table = await tableAtCursor(context, 'alignment')
    table.alignment = 'Centered'
    await context.sync()
    return 'Table centred between the margins.'
  })
}

/** Reads "4, 3, 3, 6.6" as widths in centimetres. */
export function parseWidths(text: string): number[] | null {
  const parts = text.split(/[,;\s]+/).filter(Boolean)
  const widths = parts.map(Number)
  return widths.length && widths.every((w) => Number.isFinite(w) && w >= 0.5 && w <= 30) ? widths : null
}

/** Sets the columns of the table at the cursor to widths in centimetres, first column first. */
export async function setColumnWidths(widths: number[]): Promise<string> {
  if (!widths.length || widths.some((w) => !Number.isFinite(w) || w < 0.5 || w > 30)) {
    throw new UserFacingError('Give each width in centimetres, between 0.5 and 30, separated by commas: 4, 3, 3, 6.6')
  }
  return runWord(async (context) => {
    const table = await tableAtCursor(context, 'isUniform')
    if (!table.isUniform) throw new UserFacingError(MERGED)
    const rows = table.rows
    rows.load('items')
    await context.sync()
    const cells = rows.items.map((r) => r.cells)
    cells.forEach((c) => c.load('items'))
    await context.sync()
    const columns = cells[0]?.items.length ?? 0
    for (const row of cells) row.items.forEach((cell, c) => {
      const w = widths[c]
      if (w !== undefined) cell.columnWidth = cmToPt(w)
    })
    await context.sync()
    const used = widths.slice(0, columns)
    const total = used.reduce((a, w) => a + w, 0)
    const extra = widths.length > columns ? ` The table has only ${count(columns, 'column')}, so the rest were not used.` : ''
    const fewer = used.length < columns ? ` The other ${count(columns - used.length, 'column')} kept ${columns - used.length === 1 ? 'its' : 'their'} width.` : ''
    return `Column widths set: ${listOf(used.map(formatCm))} cm (${formatCm(total)} cm in all).${fewer}${extra}`
  })
}

/**
 * Adds a total row that adds up each column of figures, or updates the totals when the last row is already a total.
 * Columns that do not add up (years, rates, percentages that are not shares, mixed units) are left blank.
 */
export async function addTotalRow(): Promise<string> {
  return runWord(async (context) => {
    const table = await tableAtCursor(context, 'values,headerRowCount,style')
    const values = table.values
    const headerRows = Math.max(1, table.headerRowCount)
    const existing = hasTotalRow(values, headerRows)
    const body = values.slice(headerRows, existing ? -1 : undefined)
    if (!body.length) throw new UserFacingError('Add rows of figures first: this table has only a header.')
    const width = body[body.length - 1]?.length ?? 0
    const header = values[headerRows - 1] ?? []
    const columns = figureColumns([...values.slice(0, headerRows), ...body], headerRows).filter((c) => c < width)
    const label = Array.from({ length: width }, (_, c) => c).find((c) => !columns.includes(c)) ?? 0
    const row: string[] = Array.from({ length: width }, (_, c) => (c === label ? 'Total' : ''))
    const summed: number[] = []
    const skipped: string[] = []
    const nameOf = (c: number): string => (header[c] ?? '').trim() || `column ${c + 1}`
    for (const c of columns.filter((c) => c !== label)) {
      const total = totalOf(body.map((r) => r[c] ?? ''), header[c] ?? '')
      if (total.ok) {
        row[c] = total.text
        summed.push(c)
      } else skipped.push(`${nameOf(c)} (${total.reason})`)
    }
    if (!summed.length) {
      const why = skipped.length ? ` ${listOf(skipped)}.` : ''
      throw new UserFacingError(`No column of figures in this table can be added up.${why}`)
    }
    const at = existing ? values.length - 1 : values.length
    if (existing) {
      // Only the totals change: the label and any figure typed in by hand stay as they are.
      for (const c of summed) table.getCell(at, c).value = row[c] ?? ''
    } else {
      const added = table.addRows('End', 1, [row])
      if (!DATA_TABLE_STYLE_NAMES.has(table.style)) added.getFirst().font.bold = true
    }
    for (const c of summed) table.getCell(at, c).horizontalAlignment = 'Right'
    table.styleTotalRow = true
    await context.sync()
    const left = skipped.length ? ` ${existing ? 'Not changed' : 'Left blank'}: ${listOf(skipped)}.` : ''
    return `${existing ? 'Total row updated' : 'Total row added'}: ${listOf(summed.map(nameOf))} added up.${left}`
  })
}

/** Deletes the rows with no text (and no pictures) from the table at the cursor. */
export async function deleteEmptyRows(): Promise<string> {
  return runWord(async (context) => {
    const table = await tableAtCursor(context, 'values')
    const values = table.values
    const candidates = values.map((r, i) => (r.every((c) => !c.trim()) ? i : -1)).filter((i) => i >= 0)
    if (!candidates.length) return 'This table has no empty rows.'
    const pictures = candidates.map((r) => (values[r] ?? []).map((_, c) => table.getCell(r, c).body.inlinePictures))
    pictures.flat().forEach((p) => p.load('items'))
    await context.sync()
    const empty = candidates.filter((_, i) => (pictures[i] ?? []).every((p) => p.items.length === 0))
    if (empty.length === values.length) throw new UserFacingError('Every row of this table is empty. Delete the whole table instead.')
    if (!empty.length) return 'This table has no empty rows: the rows without text hold pictures.'
    // From the bottom up, so the row numbers above stay right.
    for (const r of [...empty].reverse()) table.deleteRows(r, 1)
    await context.sync()
    return `${count(empty.length, 'empty row')} deleted.`
  })
}
