import { describe, expect, it } from 'vitest'
import { MINUS, autoDecimals, decimalsNeeded, formatNumber, groupThousands, unitsLabel } from '../../src/charts/format.ts'

describe('formatNumber', () => {
  it('writes thousands separators and a fixed number of decimals', () => {
    expect(formatNumber(1234567.891, { decimals: 2 })).toBe('1,234,567.89')
    expect(formatNumber(1234.5, { decimals: 0 })).toBe('1,235')
    expect(formatNumber(999.95, { decimals: 1 })).toBe('1,000.0')
    expect(formatNumber(12)).toBe('12')
  })

  it('puts an ISO currency code before the number with a space, and the scale straight after', () => {
    expect(formatNumber(250, { currency: 'USD', scale: 'm' })).toBe('USD 250m')
    expect(formatNumber(1.2, { currency: 'eur', scale: 'bn' })).toBe('EUR 1.2bn')
    expect(formatNumber(3.5, { scale: 'k' })).toBe('3.5k')
  })

  it('writes per cent with the % sign', () => {
    expect(formatNumber(12.5, { percent: true })).toBe('12.5%')
    expect(formatNumber(40, { percent: true, scale: 'm' })).toBe('40%')
  })

  it('writes negatives with a true minus sign, never brackets', () => {
    expect(formatNumber(-4)).toBe(`${MINUS}4`)
    expect(formatNumber(-250, { currency: 'USD', scale: 'm' })).toBe(`USD ${MINUS}250m`)
    expect(formatNumber(-0.001, { decimals: 1 })).toBe('0.0')
  })

  it('adds a plus sign when asked, for changes', () => {
    expect(formatNumber(12, {}, { plus: true })).toBe('+12')
    expect(formatNumber(0, {}, { plus: true })).toBe('0')
  })

  it('leaves out currency and scale for bare labels, keeping %', () => {
    expect(formatNumber(250, { currency: 'USD', scale: 'm' }, { bare: true })).toBe('250')
    expect(formatNumber(25, { percent: true }, { bare: true })).toBe('25%')
  })
})

describe('decimals', () => {
  it('finds the decimals a value needs', () => {
    expect(decimalsNeeded(12)).toBe(0)
    expect(decimalsNeeded(1.25)).toBe(2)
    expect(decimalsNeeded(0.1 + 0.2)).toBe(1)
  })

  it('chooses shared decimals for a set of labels, fewer for large numbers', () => {
    expect(autoDecimals([82.4, 96.1, 115.9])).toBe(1)
    expect(autoDecimals([1.25, 3.5])).toBe(2)
    expect(autoDecimals([120, 150])).toBe(0)
    expect(autoDecimals([33.3333, 12])).toBe(1)
    expect(autoDecimals([1234.5, 2000])).toBe(0)
    expect(autoDecimals([])).toBe(0)
  })

  it('groups thousands', () => {
    expect(groupThousands('1234567')).toBe('1,234,567')
    expect(groupThousands('999')).toBe('999')
  })

  it('names the units', () => {
    expect(unitsLabel({ currency: 'USD', scale: 'm' })).toBe('USD m')
    expect(unitsLabel({ percent: true })).toBe('%')
    expect(unitsLabel({})).toBe('')
  })
})
