import { afterEach, describe, expect, it } from 'vitest'
import { figureColumns, formatTableAtCursor, hasTotalRow, isFigure } from '../../src/word/tables.ts'
import type { FakeTable } from '../fakeWord/model.ts'
import { createModel, para } from '../fakeWord/model.ts'
import { installFakeWord, uninstallFakeWord } from '../fakeWord/install.ts'

afterEach(uninstallFakeWord)

describe('isFigure', () => {
  it.each(['1,234', '12.5%', '(300)', '-4', '\u22124.2', 'USD 250m', 'EUR 1.2bn', '3.2bn', '1.5x', '2030', '45 pp', '1\u00A0234'])('accepts %s', (v) => {
    expect(isFigure(v)).toBe(true)
  })
  it.each(['Total', 'n/a', 'Angola', 'Q1 2026', 'Phase 2', '2022 to 2025', ''])('rejects %s', (v) => {
    expect(isFigure(v)).toBe(false)
  })
})

describe('figureColumns', () => {
  const values = [
    ['Country', 'Instrument', 'Size (USD m)', 'Year', 'Share'],
    ['Angola', 'Green bond', '250', '2026', '12%'],
    ['Kenya', 'Loan', '1,200', 'n/a', '\u2013'],
    ['Total', '', '1,450', '', '100%'],
  ]
  it('finds columns whose body cells are all figures, ignoring missing-value marks', () => {
    expect(figureColumns(values, 1)).toEqual([2, 3, 4])
  })
  it('ignores a column that is empty or a mark throughout', () => {
    expect(figureColumns([['A', 'B'], ['x', ''], ['y', 'n/a']], 1)).toEqual([])
  })
})

describe('hasTotalRow', () => {
  it('spots Total, Grand total and Subtotal in the first cell of the last row', () => {
    expect(hasTotalRow([['h'], ['a'], ['Total']], 1)).toBe(true)
    expect(hasTotalRow([['h'], ['a'], ['Grand total']], 1)).toBe(true)
    expect(hasTotalRow([['h'], ['a'], ['Sub-total']], 1)).toBe(true)
    expect(hasTotalRow([['h'], ['a'], ['Totally different']], 1)).toBe(false)
    expect(hasTotalRow([['h'], ['Total']], 1)).toBe(false)
  })
})

describe('formatTableAtCursor', () => {
  const table = (): FakeTable => ({
    style: 'Table Grid',
    rows: [['Country', 'Size (USD m)'], ['Angola', '250'], ['Kenya', '1,200'], ['Total', '1,450']],
  })

  it('applies BwB Standard, repeats the header, right-aligns figures and styles the total row', async () => {
    const t = table()
    const model = createModel([para('x')], { selectionTable: t })
    model.styles.add('BwB Standard')
    model.tables.push(t)
    installFakeWord(model)
    const message = await formatTableAtCursor('standard')
    expect(t.style).toBe('BwB Standard')
    expect(t.headerRowCount).toBe(1)
    expect(t.options).toMatchObject({ styleTotalRow: true, styleFirstColumn: false, styleBandedRows: false })
    expect(t.alignments).toEqual({ '0,1': 'Right', '1,1': 'Right', '2,1': 'Right', '3,1': 'Right' })
    expect(message).toBe('BwB Standard applied, 1 column of figures right-aligned, total row styled.')
  })

  it('brings in the BwB styles first when the document lacks them', async () => {
    const t = table()
    const model = createModel([para('x')], { selectionTable: t })
    model.tables.push(t)
    installFakeWord(model)
    const message = await formatTableAtCursor('comparison')
    expect(t.style).toBe('BwB Comparison')
    expect(t.options).toMatchObject({ styleFirstColumn: true, styleTotalRow: false })
    expect(message).toMatch(/BwB styles were added/)
  })

  it('explains what to do when the cursor is not in a table', async () => {
    const model = createModel([para('x')])
    model.styles.add('BwB Standard')
    installFakeWord(model)
    await expect(formatTableAtCursor()).rejects.toThrow('Put the cursor in a table first.')
  })
})
