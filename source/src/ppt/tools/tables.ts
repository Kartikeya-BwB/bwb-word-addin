// Table tools for PowerPoint tables: the BwB format, transpose, sort, a total row, column widths, and tables from
// cells pasted from Excel. Reading a table needs PowerPointApi 1.8; formatting cells, adding rows and setting column
// widths need 1.9. Where 1.9 is missing, tools that change the table's size rebuild it in the BwB style instead.

import { CONTENT_AREA, TAGS } from '../../brand/slideGrid.ts'
import { UserFacingError } from '../../word/run.ts'
import { currentSlide } from '../render.ts'
import { requirePpt, runPpt, supportsPpt } from '../run.ts'
import type { CellLook } from './tableData.ts'
import {
  TABLE_FONT, TABLE_RULE, describeWidths, equalWidths, parseWidths, pastedGrid, planTable, sortedValues, transpose, withTotalRow,
} from './tableData.ts'

interface Selected {
  shape: PowerPoint.Shape
  table: PowerPoint.Table
  values: string[][]
}

const TOOLS = 'The table tools'
const plural = (n: number, word: string): string => `${n} ${word}${n === 1 ? '' : 's'}`

/** The selected table: a shape of type Table among the selected shapes. */
async function selectedTable(context: PowerPoint.RequestContext): Promise<Selected> {
  const shapes = context.presentation.getSelectedShapes()
  shapes.load('items/id,items/type,items/name,items/left,items/top,items/width,items/height')
  await context.sync()
  const shape = shapes.items.find((s) => s.type === 'Table')
  if (!shape) throw new UserFacingError('Select a table first: click its edge so the whole table is selected.')
  const table = shape.getTable()
  table.load('values,rowCount,columnCount')
  await context.sync()
  return { shape, table, values: table.values }
}

function rule(): PowerPoint.BorderProperties {
  return { color: `#${TABLE_RULE.colour}`, weight: TABLE_RULE.weight, dashStyle: 'Solid' }
}

/** Cell properties for a new table in the BwB look (PowerPointApi 1.8 takes them when the table is made). */
function cellProperties(values: readonly string[][]): PowerPoint.TableCellProperties[][] {
  const plan = planTable(values)
  return plan.cells.map((row) => row.map((look) => ({
    fill: { color: `#${look.fill}` },
    font: { name: TABLE_FONT.name, size: TABLE_FONT.size, color: `#${look.colour}`, bold: look.bold },
    horizontalAlignment: look.align,
    borders: { top: rule(), bottom: rule(), left: rule(), right: rule() },
  })))
}

/** Formats one existing cell (PowerPointApi 1.9). */
function formatCell(cell: PowerPoint.TableCell, look: CellLook): void {
  cell.fill.setSolidColor(`#${look.fill}`)
  cell.font.name = TABLE_FONT.name
  cell.font.size = TABLE_FONT.size
  cell.font.color = `#${look.colour}`
  cell.font.bold = look.bold
  cell.horizontalAlignment = look.align
  for (const side of [cell.borders.top, cell.borders.bottom, cell.borders.left, cell.borders.right]) {
    side.color = `#${TABLE_RULE.colour}`
    side.weight = TABLE_RULE.weight
    side.dashStyle = 'Solid'
  }
}

/** Formats every cell of a table, after taking off PowerPoint's own table style. */
function formatTable(table: PowerPoint.Table, values: readonly string[][], rows?: readonly number[]): void {
  const plan = planTable(values)
  if (!rows) table.styleSettings.style = 'NoStyleNoGrid'
  plan.cells.forEach((row, r) => {
    if (rows && !rows.includes(r)) return
    row.forEach((look, c) => formatCell(table.getCellOrNullObject(r, c), look))
  })
}

/** Puts a new BwB table where the old one was, with its name and tags, removes the old one and selects the new. */
async function rebuild(context: PowerPoint.RequestContext, old: Selected, values: string[][]): Promise<void> {
  const tags = old.shape.tags
  tags.load('items/key,items/value')
  await context.sync()
  const slide = old.shape.getParentSlide()
  const fresh = slide.shapes.addTable(values.length, values[0]?.length ?? 1, {
    values, left: old.shape.left, top: old.shape.top, width: old.shape.width, specificCellProperties: cellProperties(values),
    ...(supportsPpt('1.9') ? { style: 'NoStyleNoGrid' as const } : {}),
  })
  fresh.name = old.shape.name
  for (const t of tags.items) fresh.tags.add(t.key, t.value)
  fresh.load('id')
  old.shape.delete()
  await context.sync()
  slide.setSelectedShapes([fresh.id])
  await context.sync()
}

function onTable<T>(version: string, feature: string, fn: (context: PowerPoint.RequestContext, t: Selected) => Promise<T>): Promise<T> {
  requirePpt(version, feature)
  return runPpt(async (context) => fn(context, await selectedTable(context)))
}

/** The BwB format for the selected table: navy header, light grey rules, figures right-aligned, a total row. */
export function formatSelectedTable(): Promise<string> {
  return onTable('1.9', 'Formatting table cells', async (context, t) => {
    formatTable(t.table, t.values)
    await context.sync()
    const plan = planTable(t.values)
    const extras = [plan.figures.length ? `${plural(plan.figures.length, 'column')} of figures right-aligned` : '', plan.total ? 'the total row in bold' : '']
    return `Table formatted: navy header row, light grey rules${extras.filter(Boolean).map((e) => `, ${e}`).join('')}.`
  })
}

/** Rows become columns. The table is rebuilt in the BwB style where it stood. */
export function transposeSelectedTable(): Promise<string> {
  return onTable('1.8', TOOLS, async (context, t) => {
    const values = transpose(t.values)
    await rebuild(context, t, values)
    return `Table turned: now ${plural(values.length, 'row')} by ${plural(values[0]?.length ?? 0, 'column')}.`
  })
}

/** The header row of the selected table, for choosing a column to sort by. */
export function readTableHeaders(): Promise<string[]> {
  return onTable('1.8', TOOLS, async (_context, t) => (t.values[0] ?? []).map((h, i) => h.trim() || `Column ${i + 1}`))
}

/** Sorts the body rows by a column; the header and a total row stay put, and each cell keeps its format. */
export function sortSelectedTable(column: number, descending: boolean): Promise<string> {
  return onTable('1.8', TOOLS, async (context, t) => {
    const header = t.values[0] ?? []
    if (column < 0 || column >= header.length) throw new UserFacingError('Choose a column of this table to sort by.')
    const sorted = sortedValues(t.values, column, descending)
    sorted.values.forEach((row, r) => row.forEach((text, c) => {
      if (text !== t.values[r]?.[c]) t.table.getCellOrNullObject(r, c).text = text
    }))
    await context.sync()
    const order = sorted.numeric ? (descending ? 'largest first' : 'smallest first') : descending ? 'Z to A' : 'A to Z'
    return `Sorted by “${header[column]?.trim() || `column ${column + 1}`}”, ${order}.`
  })
}

/** Adds a total row that adds up the columns of figures, or updates the one there. */
export function addTableTotalRow(): Promise<string> {
  return onTable('1.8', TOOLS, async (context, t) => {
    const result = withTotalRow(t.values)
    if (!result.summed.length) throw new UserFacingError('There is no column of figures to add up in this table.')
    const last = result.values.length - 1
    const row = result.values[last] ?? []
    if (!result.added) {
      row.forEach((text, c) => (t.table.getCellOrNullObject(last, c).text = text))
      await context.sync()
    } else if (supportsPpt('1.9')) {
      t.table.rows.add(last, 1)
      row.forEach((text, c) => (t.table.getCellOrNullObject(last, c).text = text))
      formatTable(t.table, result.values, [last])
      await context.sync()
    } else {
      await rebuild(context, t, result.values)
    }
    const left = result.skipped.length ? ` Left blank: ${result.skipped.join('; ')}.` : ''
    return `${result.added ? 'Total row added' : 'Total row updated'}: ${plural(result.summed.length, 'column')} added up.${left}`
  })
}

/** Gives every column the same width, keeping the table's width. */
export function equaliseColumnWidths(): Promise<string> {
  return onTable('1.9', 'Setting column widths', async (context, t) => {
    const widths = equalWidths(t.shape.width, t.values[0]?.length ?? 0)
    widths.forEach((w, i) => (t.table.columns.getItemAt(i).width = w))
    await context.sync()
    return `Columns set to the same width: ${describeWidths(widths.slice(0, 1))} each.`
  })
}

/** Sets the column widths from a list in centimetres. */
export function setColumnWidths(text: string): Promise<string> {
  return onTable('1.9', 'Setting column widths', async (context, t) => {
    const widths = parseWidths(text, t.values[0]?.length ?? 0)
    widths.forEach((w, i) => (t.table.columns.getItemAt(i).width = w))
    await context.sync()
    const total = widths.reduce((a, b) => a + b, 0)
    return `Column widths set: ${describeWidths(widths)} (${describeWidths([total])} in all).`
  })
}

/** A new BwB table in the content area of the current slide, from cells copied from Excel. */
export async function pasteNewTable(text: string): Promise<string> {
  requirePpt('1.8', 'Making a table')
  const values = pastedGrid(text)
  return runPpt(async (context) => {
    const slide = await currentSlide(context)
    const table = slide.shapes.addTable(values.length, values[0]?.length ?? 1, {
      values, left: CONTENT_AREA.left, top: CONTENT_AREA.top, width: CONTENT_AREA.width, specificCellProperties: cellProperties(values),
      ...(supportsPpt('1.9') ? { style: 'NoStyleNoGrid' as const } : {}),
    })
    table.name = 'BwB table'
    table.tags.add(TAGS.element, 'table')
    table.load('id')
    await context.sync()
    slide.setSelectedShapes([table.id])
    await context.sync()
    return `Table of ${plural(values.length, 'row')} by ${plural(values[0]?.length ?? 0, 'column')} added in the BwB style.`
  })
}

/** Pasted cells into the selected table: in place when the size matches, otherwise rebuilt to the new size. */
export async function pasteIntoSelectedTable(text: string): Promise<string> {
  const values = pastedGrid(text)
  return onTable('1.8', TOOLS, async (context, t) => {
    const sameSize = values.length === t.values.length && values[0]?.length === t.values[0]?.length
    if (sameSize) {
      values.forEach((row, r) => row.forEach((cell, c) => (t.table.getCellOrNullObject(r, c).text = cell)))
      await context.sync()
      return 'Values replaced. The table kept its format.'
    }
    await rebuild(context, t, values)
    return `The table was rebuilt where it stood, at ${plural(values.length, 'row')} by ${plural(values[0]?.length ?? 0, 'column')}, in the BwB style.`
  })
}
