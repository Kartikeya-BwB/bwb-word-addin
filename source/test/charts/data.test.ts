import { describe, expect, it } from 'vitest'
import { parseNumber, parsePasted, toSeriesTable, unitsHint } from '../../src/charts/data.ts'

describe('parseNumber', () => {
  it.each([
    ['1,234.5', 1234.5],
    ['1234.5', 1234.5],
    ['(300)', -300],
    ['-4', -4],
    ['\u{2212}4', -4],
    ['+7', 7],
    ['0.25', 0.25],
    ['.5', 0.5],
    ['12%', 12],
    ['12 %', 12],
    ['(12%)', -12],
    ['5m', 5],
    ['3.5k', 3.5],
    ['1.2bn', 1.2],
    ['USD 1.2bn', 1.2],
    ['USD1.2bn', 1.2],
    ['USD -250m', -250],
    ['(USD 250m)', -250],
    ['-USD 250m', -250],
    ['250m USD', 250],
    ['$5m', 5],
    ['\u{00A3}3.4bn', 3.4],
    ['\u{20AC}12', 12],
    ['1 234 567', 1234567],
    ['2025', 2025],
  ])('reads %s as %s', (text, value) => {
    const parsed = parseNumber(text)
    expect(parsed.ok).toBe(true)
    expect(parsed.value).toBeCloseTo(value, 9)
  })

  it('reports the units each cell carries', () => {
    expect(parseNumber('USD 1.2bn')).toMatchObject({ currency: 'USD', scale: 'bn' })
    expect(parseNumber('$5m')).toMatchObject({ currency: 'USD', scale: 'm' })
    expect(parseNumber('\u{00A3}3.4bn')).toMatchObject({ currency: 'GBP', scale: 'bn' })
    expect(parseNumber('\u{20AC}12')).toMatchObject({ currency: 'EUR' })
    expect(parseNumber('12%')).toMatchObject({ percent: true })
    expect(parseNumber('3.5k')).toMatchObject({ scale: 'k' })
    expect(parseNumber('5 million')).toMatchObject({ value: 5, scale: 'm' })
    expect(parseNumber('2B')).toMatchObject({ value: 2, scale: 'bn' })
    expect(parseNumber('1,234').currency).toBeUndefined()
  })

  it.each(['', '  ', 'n/a', 'N/A', 'na', '-', '\u{2013}', '..', '#N/A', '#DIV/0!'])('treats %j as a gap', (text) => {
    expect(parseNumber(text)).toEqual({ value: null, ok: true })
  })

  it('flags text that is not a number', () => {
    expect(parseNumber('about ten')).toEqual({ value: null, ok: false })
    expect(parseNumber('Q1')).toEqual({ value: null, ok: false })
    expect(parseNumber('12..5')).toEqual({ value: null, ok: false })
  })

  it('reads decimal commas', () => {
    expect(parseNumber('1.234,5', { decimalComma: true }).value).toBeCloseTo(1234.5)
    expect(parseNumber('1,5', { decimalComma: true }).value).toBeCloseTo(1.5)
    expect(parseNumber('1.234,5').value).toBeCloseTo(1234.5)
    expect(parseNumber('1,5').value).toBeCloseTo(1.5)
    expect(parseNumber('1,234').value).toBe(1234)
  })
})

describe('parsePasted', () => {
  it('splits what Excel copies: tabs, CRLF line ends and a trailing newline', () => {
    const table = parsePasted('Year\tFunding\r\n2022\t120\r\n2023\t150\r\n')
    expect(table.delimiter).toBe('\t')
    expect(table.grid).toEqual([['Year', 'Funding'], ['2022', '120'], ['2023', '150']])
  })

  it('keeps quoted cells whole, with tabs, newlines and doubled quotes inside', () => {
    const table = parsePasted('Name\tNote\n"Line one\nline two"\t"He said ""yes"""')
    expect(table.grid).toEqual([['Name', 'Note'], ['Line one\nline two', 'He said "yes"']])
  })

  it('falls back to semicolons, with European decimals', () => {
    const table = parsePasted('Region;Share\nNorth;12,5\nSouth;1.234,5')
    expect(table.delimiter).toBe(';')
    expect(table.decimalComma).toBe(true)
    const series = toSeriesTable(table.grid, { decimalComma: table.decimalComma })
    expect(series.series[0]?.values).toEqual([12.5, 1234.5])
  })

  it('falls back to commas, honouring quotes round numbers with separators', () => {
    const table = parsePasted('Year,Funding\n2022,"1,234"\n2023,150')
    expect(table.delimiter).toBe(',')
    expect(table.grid[1]).toEqual(['2022', '1,234'])
  })

  it('pads ragged rows and drops empty trailing rows and columns', () => {
    const table = parsePasted('A\tB\tC\t\n1\t2\n\n')
    expect(table.grid).toEqual([['A', 'B', 'C'], ['1', '2', '']])
  })

  it('reads a single column of lines', () => {
    expect(parsePasted('one\ntwo').grid).toEqual([['one'], ['two']])
  })
})

describe('toSeriesTable', () => {
  it('takes series from the first row and categories from the first column', () => {
    const table = toSeriesTable([['Year', 'Grants', 'Loans'], ['2022', '1', '2'], ['2023', '', 'n/a'], ['2024', '3', 'x']])
    expect(table.corner).toBe('Year')
    expect(table.categories).toEqual(['2022', '2023', '2024'])
    expect(table.series.map((s) => s.name)).toEqual(['Grants', 'Loans'])
    expect(table.series[0]?.values).toEqual([1, null, 3])
    expect(table.series[1]?.values).toEqual([2, null, null])
    expect(table.invalid).toEqual([{ row: 3, column: 2, text: 'x' }])
  })

  it('puts mixed scales into the scale most cells use', () => {
    const table = toSeriesTable([['Year', 'Size'], ['2023', 'USD 800m'], ['2024', 'USD 1.2bn'], ['2025', 'USD 1.5bn']])
    expect(table.hint).toEqual({ currency: 'USD', scale: 'bn' })
    expect(table.series[0]?.values.map((v) => Number(v?.toFixed(6)))).toEqual([0.8, 1.2, 1.5])
  })

  it('names unnamed series and categories', () => {
    const table = toSeriesTable([['', ''], ['', '5']])
    expect(table.series[0]?.name).toBe('Series 1')
    expect(table.categories).toEqual(['1'])
  })
})

describe('unitsHint', () => {
  it('detects a currency all cells agree on, the usual scale, and per cent', () => {
    expect(unitsHint([parseNumber('USD 5m'), parseNumber('USD 7m'), parseNumber('3')])).toEqual({ currency: 'USD', scale: 'm' })
    expect(unitsHint([parseNumber('USD 5m'), parseNumber('EUR 7m')])).toEqual({ scale: 'm' })
    expect(unitsHint([parseNumber('12%'), parseNumber('30%'), parseNumber('')])).toEqual({ percent: true })
    expect(unitsHint([parseNumber('12'), parseNumber('30')])).toEqual({})
  })
})
