import { afterEach, describe, expect, it } from 'vitest'
import { readTableColumns, sortTable, tableToText, textToTable, transposeTable } from '../../../src/word/tools/tableData.ts'
import {
  addTotalRow, centreTable, deleteEmptyRows, distributeColumns, fitTableToContents, fitTableToWindow, parseWidths, setColumnWidths,
} from '../../../src/word/tools/tableTools.ts'
import { CM_TO_PT } from '../../../src/word/tools/common.ts'
import type { FakeModel, FakeTable } from '../../fakeWord/model.ts'
import { createModel, para } from '../../fakeWord/model.ts'
import { installFakeWord, uninstallFakeWord } from '../../fakeWord/install.ts'

afterEach(uninstallFakeWord)

/** A document with the cursor in a table whose cells are paragraphs, followed by a caption. */
function withTable(rows: string[][], style = 'BwB Standard', over: Partial<FakeModel> = {}): { model: FakeModel; table: FakeTable } {
  const table: FakeTable = { style, rows: rows.map((r) => [...r]), headerRowCount: 1 }
  const cells = rows.flatMap((r) => r.map((t) => para(t, 'Normal', { tableNestingLevel: 1, table })))
  const model = createModel([para('Intro'), ...cells, para('Table 1. Green bonds', 'BwB Table Caption'), para('After')], {
    selectionTable: table, ...over,
  })
  model.tables.push(table)
  model.styles.add(style)
  return { model, table }
}

const SIZES = [['Country', 'Size (USD m)', 'Year', 'Share'], ['Angola', '250', '2026', '20%'], ['Kenya', '1,000', '2025', '80%']]

describe('table layout', () => {
  it('evens out the columns, but not of a table with merged cells', async () => {
    const { model, table } = withTable(SIZES)
    installFakeWord(model)
    expect(await distributeColumns()).toBe('Columns made the same width.')
    expect(table.actions).toEqual(['distributeColumns'])
    table.rows[0] = ['Merged header']
    await expect(distributeColumns()).rejects.toThrow(/merged cells/)
  })

  it('fits the table to the window, centres it and explains when the cursor is not in a table', async () => {
    const { model, table } = withTable(SIZES)
    installFakeWord(model)
    await fitTableToWindow()
    await centreTable()
    expect(table.actions).toEqual(['autoFitWindow'])
    expect(table.alignment).toBe('Centered')
    delete model.selectionTable
    await expect(centreTable()).rejects.toThrow('Put the cursor in a table first.')
  })

  it('fits columns to their contents on the desktop, and says why not on the web', async () => {
    const { model, table } = withTable(SIZES, 'BwB Standard', { desktopApiVersion: 1.4 })
    installFakeWord(model)
    expect(await fitTableToContents()).toBe('Columns fitted to their contents.')
    expect(table.actions).toEqual(['autoFitBehavior:Content'])
    model.desktopApiVersion = 0
    model.platform = 'OfficeOnline'
    await expect(fitTableToContents()).rejects.toThrow('Needs Word for Windows or Mac.')
  })

  it('sets column widths in centimetres', async () => {
    const { model, table } = withTable(SIZES)
    installFakeWord(model)
    expect(await setColumnWidths([4, 3.5])).toBe('Column widths set: 4 and 3.5 cm (7.5 cm in all). The other 2 columns kept their width.')
    expect(Math.round((table.widths?.['2,1'] ?? 0) / CM_TO_PT * 10) / 10).toBe(3.5)
    expect(table.widths?.['0,2']).toBeUndefined()
    expect(parseWidths('4, 3; 3 6.6')).toEqual([4, 3, 3, 6.6])
    expect(parseWidths('4, wide')).toBeNull()
    await expect(setColumnWidths([0.1])).rejects.toThrow(/between 0.5 and 30/)
  })
})

describe('addTotalRow', () => {
  it('adds a total row that adds up each column of figures, leaving years blank', async () => {
    const { model, table } = withTable(SIZES)
    installFakeWord(model)
    expect(await addTotalRow()).toBe('Total row added: Size (USD m) and Share added up. Left blank: Year (years).')
    expect(table.rows[3]).toEqual(['Total', '1,250', '', '100%'])
    expect(table.alignments).toEqual({ '3,1': 'Right', '3,3': 'Right' })
    expect(table.options).toMatchObject({ styleTotalRow: true })
    expect(table.boldRows).toBeUndefined()
  })

  it('updates the totals of an existing total row, keeping what was typed in the other cells', async () => {
    const { model, table } = withTable([...SIZES, ['Grand total', '9', 'n/a', '100%']])
    installFakeWord(model)
    expect(await addTotalRow()).toBe('Total row updated: Size (USD m) and Share added up. Not changed: Year (years).')
    expect(table.rows).toHaveLength(4)
    expect(table.rows[3]).toEqual(['Grand total', '1,250', 'n/a', '100%'])
  })

  it('makes the new row bold in a table without a BwB style', async () => {
    const { model, table } = withTable(SIZES, 'Table Grid')
    installFakeWord(model)
    await addTotalRow()
    expect(table.boldRows).toEqual([3])
  })

  it('explains when nothing can be added up', async () => {
    const { model } = withTable([['Country', 'Year'], ['Angola', '2026'], ['Kenya', '2025']])
    installFakeWord(model)
    await expect(addTotalRow()).rejects.toThrow('No column of figures in this table can be added up. Year (years).')
  })
})

describe('deleteEmptyRows', () => {
  it('deletes rows with no text, keeping rows that hold a picture', async () => {
    const { model, table } = withTable([['A', 'B'], ['', ' '], ['1', '2'], ['', ''], ['', '']])
    table.pictures = { '3,0': 1 }
    installFakeWord(model)
    expect(await deleteEmptyRows()).toBe('2 empty rows deleted.')
    expect(table.rows).toEqual([['A', 'B'], ['1', '2'], ['', '']])
    table.pictures = {}
    table.rows = [['', '']]
    await expect(deleteEmptyRows()).rejects.toThrow(/Every row of this table is empty/)
  })

  it('says so when there are no empty rows', async () => {
    const { model } = withTable(SIZES)
    installFakeWord(model)
    expect(await deleteEmptyRows()).toBe('This table has no empty rows.')
  })
})

describe('sorting and transposing', () => {
  it('reads the column headings and sorts by value, keeping the header and the total row', async () => {
    const { model, table } = withTable([['Country', 'Size'], ['Kenya', '1,200'], ['Angola', '250'], ['Benin', '600'], ['Total', '2,050']])
    installFakeWord(model)
    expect(await readTableColumns()).toEqual(['Country', 'Size'])
    expect(await sortTable({ column: 1, descending: true, numeric: true })).toBe('Rows sorted by Size, largest first.')
    expect(table.rows.map((r) => r[0])).toEqual(['Country', 'Kenya', 'Benin', 'Angola', 'Total'])
    expect(await sortTable({ column: 0, descending: false, numeric: false })).toBe('Rows sorted by Country, A to Z.')
    expect(table.rows.map((r) => r[0])).toEqual(['Country', 'Angola', 'Benin', 'Kenya', 'Total'])
    expect(await sortTable({ column: 0, descending: false, numeric: false })).toBe('The rows were already in order by Country, A to Z.')
    await expect(sortTable({ column: 5, descending: false, numeric: false })).rejects.toThrow('This table has 2 columns: choose one of them.')
  })

  it('swaps rows and columns, keeping the style and the caption below', async () => {
    const { model } = withTable([['Country', 'Size'], ['Kenya', '1,200'], ['Angola', '250']])
    installFakeWord(model)
    expect(await transposeTable()).toBe('Rows and columns swapped: the table is now 2 rows by 3 columns. Check the column widths.')
    expect(model.tables).toHaveLength(1)
    const fresh = model.tables[0]
    expect(fresh?.rows).toEqual([['Country', 'Kenya', 'Angola'], ['Size', '1,200', '250']])
    expect(fresh?.style).toBe('BwB Standard')
    expect(fresh?.alignments).toEqual({ '0,1': 'Right', '1,1': 'Right', '0,2': 'Right', '1,2': 'Right' })
    const texts = model.paragraphs.map((p) => p.text)
    expect(texts[0]).toBe('Intro')
    expect(texts.slice(-2)).toEqual(['Table 1. Green bonds', 'After'])
    expect(model.paragraphs.filter((p) => p.table === fresh)).toHaveLength(6)
  })

  it('turns a table into tab-separated lines', async () => {
    const { model } = withTable([['Country', 'Size'], ['Kenya', '1,200']])
    installFakeWord(model)
    expect(await tableToText()).toBe('Table turned into 2 lines of text, with tabs between the cells.')
    expect(model.tables).toHaveLength(0)
    expect(model.paragraphs.map((p) => p.text)).toEqual(['Intro', 'Country\tSize', 'Kenya\t1,200', 'Table 1. Green bonds', 'After'])
    expect(model.paragraphs[1]?.styleBuiltIn).toBe('Normal')
  })
})

describe('textToTable', () => {
  it('turns tab-separated lines into a BwB table with a caption, figures right-aligned', async () => {
    const lines = [para('Country\tSize (USD m)'), para('Kenya\t1,200'), para('Angola\t250'), para('Total\t1,450')]
    const model = createModel([para('Intro'), ...lines, para('After')], { selection: lines })
    installFakeWord(model)
    const message = await textToTable()
    expect(message).toBe('Table of 3 rows and 2 columns made from the text, split on tabs. Type over the caption placeholder.')
    expect(model.inserts[0]?.location).toBe('Replace')
    expect(model.tables[0]?.rows).toEqual([['Country', 'Size (USD m)'], ['Kenya', '1,200'], ['Angola', '250'], ['Total', '1,450']])
    const texts = model.paragraphs.map((p) => p.text)
    expect(texts[0]).toBe('Intro')
    expect(texts[texts.length - 1]).toBe('After')
    expect(texts.some((t) => /^Table 1\./.test(t))).toBe(true)
    expect(model.paragraphs.find((p) => p.text === '1,200')?.style).toMatch(/Number/)
    expect(model.paragraphs.find((p) => p.text === '1,450')?.style).toMatch(/Total/)
    expect(texts.includes('Kenya\t1,200')).toBe(false)
  })

  it('splits on semicolons or commas when there are no tabs, escaping the builder\u2019s mark-up', async () => {
    const lines = [para('Name; Note'), para('Fund *A*; [draft]')]
    const model = createModel([...lines, para('')], { selection: lines })
    installFakeWord(model)
    await textToTable()
    expect(model.tables[0]?.rows).toEqual([['Name', 'Note'], ['Fund *A*', '[draft]']])
  })

  it('explains what it needs', async () => {
    const plain = [para('Just text'), para('More text')]
    installFakeWord(createModel(plain, { selection: plain }))
    await expect(textToTable()).rejects.toThrow('Separate the columns with tabs, semicolons or commas, then select the lines.')
    const one = [para('A\tB')]
    installFakeWord(createModel(one, { selection: one }))
    await expect(textToTable()).rejects.toThrow('Select at least two lines: the column headings and one row.')
    const table: FakeTable = { style: 'BwB Standard', rows: [['a\tb']] }
    const cell = [para('a\tb', 'Normal', { tableNestingLevel: 1, table })]
    installFakeWord(createModel(cell, { selection: cell }))
    await expect(textToTable()).rejects.toThrow('Select lines of text outside a table.')
  })
})
