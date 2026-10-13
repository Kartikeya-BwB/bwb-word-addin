import { describe, expect, it } from 'vitest'
import { parseFigure, rowsFromLines, sortRows, totalOf, transpose } from '../../../src/word/tools/figures.ts'

describe('parseFigure', () => {
  it.each([
    ['1,234', 1234, 0],
    ['12.5%', 12.5, 1],
    ['(300)', -300, 0],
    ['-4', -4, 0],
    ['\u22124.2', -4.2, 1],
    ['USD 250m', 250e6, 0],
    ['EUR 1.2bn', 1.2e9, 1],
    ['3.2bn', 3.2e9, 1],
    ['1\u00a0234', 1234, 0],
    ['45 pp', 45, 0],
    ['1.5x', 1.5, 1],
    ['+7', 7, 0],
    ['USD -5m', -5e6, 0],
    ['12k', 12000, 0],
    ['250*', 250, 0],
  ])('reads %s', (text, value, decimals) => {
    const f = parseFigure(text)
    expect(f?.value).toBeCloseTo(value)
    expect(f?.decimals).toBe(decimals)
  })

  it('keeps the currency code, scale, suffix and the way negatives are written', () => {
    expect(parseFigure('usd 1.2BN')).toMatchObject({ currency: 'USD', scale: 'bn', suffix: '' })
    expect(parseFigure('(300)')).toMatchObject({ parentheses: true, minus: '' })
    expect(parseFigure('\u22124')).toMatchObject({ parentheses: false, minus: '\u2212' })
    expect(parseFigure('12.5%')).toMatchObject({ suffix: '%', scale: '' })
  })

  it.each(['Total', 'n/a', 'Kenya', '1,5', '12,50', '1.2.3', '', 'Phase 2'])('rejects %j', (text) => {
    expect(parseFigure(text)).toBeNull()
  })
})

describe('totalOf', () => {
  it('adds up plain figures with thousands separators', () => {
    expect(totalOf(['250', '1,200', '-'])).toEqual({ ok: true, text: '1,450', value: 1450 })
  })

  it('adds currency amounts in mixed scales, writing the total in the largest', () => {
    expect(totalOf(['USD 250m', 'USD 1.2bn'])).toMatchObject({ ok: true, text: 'USD 1.45bn' })
    expect(totalOf(['USD 250m', 'USD 300m'])).toMatchObject({ ok: true, text: 'USD 550m' })
  })

  it('writes a negative total the way the column writes negatives', () => {
    expect(totalOf(['(300)', '100'])).toMatchObject({ ok: true, text: '(200)' })
    expect(totalOf(['-300', '100'])).toMatchObject({ ok: true, text: '-200' })
    expect(totalOf(['(300)', '500'])).toMatchObject({ ok: true, text: '200' })
  })

  it('totals shares of a whole, keeping their decimal places', () => {
    expect(totalOf(['12.5%', '87.5%'])).toMatchObject({ ok: true, text: '100.0%' })
    expect(totalOf(['40%', '60%'])).toMatchObject({ ok: true, text: '100%' })
  })

  it('refuses columns that do not add up, saying why', () => {
    expect(totalOf(['5%', '7%'])).toEqual({ ok: false, reason: 'percentages that are not shares of a whole' })
    expect(totalOf(['2024', '2025'])).toEqual({ ok: false, reason: 'years' })
    expect(totalOf(['USD 5m', 'EUR 4m'])).toEqual({ ok: false, reason: 'mixed currencies' })
    expect(totalOf(['5m', '300'])).toEqual({ ok: false, reason: 'mixed units' })
    expect(totalOf(['1.5x', '2x'])).toEqual({ ok: false, reason: 'multiples or percentage points' })
    expect(totalOf(['5', 'Kenya'])).toEqual({ ok: false, reason: 'text among the figures' })
    expect(totalOf(['n/a', '-'])).toEqual({ ok: false, reason: 'no figures' })
    expect(totalOf(['3.2', '4.1'], 'Interest rate (%)')).toEqual({ ok: false, reason: 'does not add up' })
    expect(totalOf(['1', '2'], 'No.')).toEqual({ ok: false, reason: 'does not add up' })
    expect(totalOf(['2024', '2025'], 'Year')).toEqual({ ok: false, reason: 'years' })
  })

  it('keeps enough decimal places to show the sum exactly', () => {
    expect(totalOf(['0.25', '1.1'])).toMatchObject({ ok: true, text: '1.35' })
    expect(totalOf(['12.50', '3'])).toMatchObject({ ok: true, text: '15.50' })
  })
})

describe('sortRows', () => {
  const values = [
    ['Country', 'Size (USD m)'],
    ['Kenya', '1,200'],
    ['angola', '250'],
    ['Benin', 'n/a'],
    ['Total', '1,450'],
  ]

  it('sorts the body alphabetically, ignoring case, keeping the header and the total row', () => {
    const sorted = sortRows(values, { column: 0, descending: false, numeric: false, headerRows: 1, keepLast: true })
    expect(sorted.map((r) => r[0])).toEqual(['Country', 'angola', 'Benin', 'Kenya', 'Total'])
  })

  it('sorts by value, largest first, with cells that are not figures last', () => {
    const sorted = sortRows(values, { column: 1, descending: true, numeric: true, headerRows: 1, keepLast: true })
    expect(sorted.map((r) => r[0])).toEqual(['Country', 'Kenya', 'angola', 'Benin', 'Total'])
  })

  it('sorts numbers inside text naturally and keeps equal rows in their order', () => {
    const rows = [['Phase'], ['Phase 10'], ['Phase 2'], ['Phase 2'], ['']]
    const sorted = sortRows(rows, { column: 0, descending: false, numeric: false, headerRows: 1, keepLast: false })
    expect(sorted.map((r) => r[0])).toEqual(['Phase', 'Phase 2', 'Phase 2', 'Phase 10', ''])
    expect(sorted[1]).toBe(rows[2])
  })
})

describe('transpose', () => {
  it('turns rows into columns, padding short rows', () => {
    expect(transpose([['A', 'B', 'C'], ['1', '2']])).toEqual([['A', '1'], ['B', '2'], ['C', '']])
  })
})

describe('rowsFromLines', () => {
  it('splits on tabs first', () => {
    expect(rowsFromLines(['Country\tSize', 'Kenya\t1,200', '', 'Angola\t250'])).toEqual({
      delimiter: 'tab', rows: [['Country', 'Size'], ['Kenya', '1,200'], ['Angola', '250']],
    })
  })

  it('then on semicolons, padding short rows', () => {
    expect(rowsFromLines(['A; B; C', '1; 2'])).toEqual({ delimiter: 'semicolon', rows: [['A', 'B', 'C'], ['1', '2', '']] })
  })

  it('then on commas, keeping thousands separators and quoted text together', () => {
    expect(rowsFromLines(['Country, Size, Year', 'Kenya, 1,200, 2025', '"Congo, Republic of",250,2024'])?.rows).toEqual([
      ['Country', 'Size', 'Year'],
      ['Kenya', '1,200', '2025'],
      ['Congo, Republic of', '250', '2024'],
    ])
  })

  it('drops empty trailing columns, and gives null when there is nothing to split on', () => {
    expect(rowsFromLines(['A\tB\t', 'C\tD\t'])?.rows).toEqual([['A', 'B'], ['C', 'D']])
    expect(rowsFromLines(['Just a line', 'and another'])).toBeNull()
    expect(rowsFromLines(['Only\t', 'one\t'])).toBeNull()
  })
})
