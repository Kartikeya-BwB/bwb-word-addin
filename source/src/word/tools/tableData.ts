// Moving table data: sorting rows, swapping rows and columns, turning lines of text into a BwB table and back.

import { DATA_TABLE_STYLE_NAMES } from '../../brand/styles.ts'
import type { Block } from '../docxApi.ts'
import { insertBlocks } from '../insert.ts'
import { UserFacingError, runWord } from '../run.ts'
import { figureColumns, hasTotalRow } from '../tables.ts'
import { count, escapeMarkup } from './common.ts'
import { rowsFromLines, sortRows, transpose } from './figures.ts'
import { MERGED, tableAtCursor } from './tableTools.ts'

/** The column headings of the table at the cursor, for choosing a column to sort by. */
export async function readTableColumns(): Promise<string[]> {
  return runWord(async (context) => {
    const table = await tableAtCursor(context, 'values,headerRowCount')
    const header = table.values[Math.max(1, table.headerRowCount) - 1] ?? []
    return header.map((h, i) => h.trim() || `Column ${i + 1}`)
  })
}

export interface SortChoice {
  /** 0 for the first column. */
  column: number
  descending: boolean
  numeric: boolean
}

/** Sorts the body rows of the table at the cursor. The header rows and a total row stay where they are. */
export async function sortTable(choice: SortChoice): Promise<string> {
  return runWord(async (context) => {
    const table = await tableAtCursor(context, 'values,headerRowCount,isUniform')
    if (!table.isUniform) throw new UserFacingError(MERGED)
    const values = table.values
    const headerRows = Math.max(1, table.headerRowCount)
    const width = values[0]?.length ?? 0
    if (choice.column < 0 || choice.column >= width) throw new UserFacingError(`This table has ${count(width, 'column')}: choose one of them.`)
    if (values.length - headerRows < 2) throw new UserFacingError('There is nothing to sort: the table has fewer than two rows below the header.')
    const sorted = sortRows(values, { ...choice, headerRows, keepLast: hasTotalRow(values, headerRows) })
    let changed = 0
    sorted.forEach((row, r) =>
      row.forEach((v, c) => {
        if (values[r]?.[c] !== v) {
          table.getCell(r, c).value = v
          changed += 1
        }
      }),
    )
    await context.sync()
    const name = (values[headerRows - 1]?.[choice.column] ?? '').trim() || `column ${choice.column + 1}`
    const order = choice.numeric ? (choice.descending ? 'largest first' : 'smallest first') : choice.descending ? 'Z to A' : 'A to Z'
    return changed ? `Rows sorted by ${name}, ${order}.` : `The rows were already in order by ${name}, ${order}.`
  })
}

/** Swaps the rows and columns of the table at the cursor, rebuilding it in the same table style. */
export async function transposeTable(): Promise<string> {
  return runWord(async (context) => {
    const table = await tableAtCursor(context, 'values,headerRowCount,isUniform,style,alignment')
    if (!table.isUniform) throw new UserFacingError(MERGED)
    const values = table.values
    const swapped = transpose(values)
    const rows = swapped.length
    const columns = swapped[0]?.length ?? 0
    if (!rows || !columns) throw new UserFacingError('This table is empty.')
    // A paragraph between the old and new table stops Word joining them; both go once the new one is in.
    const spacer = table.insertParagraph('', 'After')
    const fresh = spacer.insertTable(rows, columns, 'After', swapped)
    if (table.style) fresh.style = table.style
    fresh.headerRowCount = 1
    const alignment = table.alignment
    if (alignment === 'Left' || alignment === 'Centered' || alignment === 'Right') fresh.alignment = alignment
    if (DATA_TABLE_STYLE_NAMES.has(table.style)) {
      for (const c of figureColumns(swapped, 1)) {
        for (let r = 0; r < rows; r += 1) fresh.getCell(r, c).horizontalAlignment = 'Right'
      }
    }
    table.delete()
    spacer.delete()
    fresh.select('Start')
    await context.sync()
    return `Rows and columns swapped: the table is now ${count(rows, 'row')} by ${count(columns, 'column')}. Check the column widths.`
  })
}

/** Turns the table at the cursor into lines of text, its cells separated by tabs. */
export async function tableToText(): Promise<string> {
  return runWord(async (context) => {
    const table = await tableAtCursor(context, 'values')
    const lines = table.values.map((row) => row.map((c) => c.replace(/[\r\n\u000b]+/g, ' ').trim()).join('\t'))
    let last: Word.Paragraph | null = null
    for (const line of lines) {
      last = last ? last.insertParagraph(line, 'After') : table.insertParagraph(line, 'After')
      last.styleBuiltIn = 'Normal'
    }
    table.delete()
    await context.sync()
    return `Table turned into ${count(lines.length, 'line')} of text, with tabs between the cells.`
  })
}

const DELIMITERS = { tab: 'tabs', semicolon: 'semicolons', comma: 'commas' } as const

/**
 * Turns the selected lines into a BwB table: split on tabs, else semicolons, else commas. The first line becomes the
 * header row, columns of figures are right-aligned, a last line starting "Total" becomes the total row, and the
 * table gets a caption to fill in.
 */
export async function textToTable(): Promise<string> {
  const parsed = await runWord(async (context) => {
    const paragraphs = context.document.getSelection().paragraphs
    paragraphs.load('items/text,items/tableNestingLevel')
    await context.sync()
    const items = paragraphs.items
    if (items.some((p) => p.tableNestingLevel > 0)) throw new UserFacingError('Select lines of text outside a table.')
    const result = rowsFromLines(items.flatMap((p) => p.text.split('\u000b')))
    if (!result) throw new UserFacingError('Separate the columns with tabs, semicolons or commas, then select the lines.')
    if (result.rows.length < 2) throw new UserFacingError('Select at least two lines: the column headings and one row.')
    const first = items[0]
    const last = items[items.length - 1]
    if (!first || !last) throw new UserFacingError('Select the lines to turn into a table.')
    // Clear the lines down to one empty paragraph with the cursor in it: insertBlocks puts the table there.
    const lines = first.getRange('Start').expandTo(last.getRange('Content'))
    lines.delete()
    lines.select('Start')
    await context.sync()
    return result
  })
  const [header = [], ...rows] = parsed.rows
  const figures = figureColumns(parsed.rows, 1)
  const block: Block = {
    type: 'table',
    header: header.map(escapeMarkup),
    rows: rows.map((r) => r.map(escapeMarkup)),
    align: header.map((_, c) => (figures.includes(c) ? 'right' : null)),
    total_row: hasTotalRow(parsed.rows, 1),
    caption: '{field:Table title}',
  }
  const { warnings } = await insertBlocks([block])
  const done = `Table of ${count(rows.length, 'row')} and ${count(header.length, 'column')} made from the text, split on ${DELIMITERS[parsed.delimiter]}. Type over the caption placeholder.`
  return [done, ...warnings].join(' ')
}
