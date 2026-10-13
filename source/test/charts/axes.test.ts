import { describe, expect, it } from 'vitest'
import { niceScale, niceStep, project, stepDecimals } from '../../src/charts/axes.ts'

const isNice = (step: number): boolean => {
  const mantissa = step / 10 ** Math.floor(Math.log10(step))
  return [1, 2, 2.5, 5].some((m) => Math.abs(mantissa - m) < 1e-9)
}

describe('niceStep', () => {
  it.each([
    [0.7, 1],
    [1, 1],
    [1.1, 2],
    [2.2, 2.5],
    [3, 5],
    [7, 10],
    [23, 25],
    [26, 50],
    [0.013, 0.02],
    [180, 200],
    [1200, 2000],
  ])('rounds %s up to %s', (raw, step) => {
    expect(niceStep(raw)).toBeCloseTo(step, 12)
  })

  it('falls back to 1 for nonsense', () => {
    expect(niceStep(0)).toBe(1)
    expect(niceStep(Number.NaN)).toBe(1)
  })
})

describe('niceScale', () => {
  it('covers the data with round ticks from zero', () => {
    const scale = niceScale(0, 240, { maxTicks: 6 })
    expect(scale).toEqual({ min: 0, max: 250, step: 50, ticks: [0, 50, 100, 150, 200, 250] })
  })

  it('always uses steps of 1, 2, 2.5 or 5 times a power of ten', () => {
    for (const hi of [3, 7.7, 13, 99, 101, 333, 4321, 0.087]) {
      for (const maxTicks of [3, 4, 5, 6, 8]) {
        const scale = niceScale(0, hi, { maxTicks })
        expect(isNice(scale.step)).toBe(true)
        expect(scale.ticks.length).toBeLessThanOrEqual(maxTicks)
        expect(scale.max).toBeGreaterThanOrEqual(hi)
        expect(scale.ticks[0]).toBe(0)
      }
    }
  })

  it('spans negative values', () => {
    const scale = niceScale(-35, 80, { maxTicks: 6 })
    expect(scale.min).toBeLessThanOrEqual(-35)
    expect(scale.max).toBeGreaterThanOrEqual(80)
    expect(scale.ticks).toContain(0)
  })

  it('can leave zero out for lines far from it', () => {
    const scale = niceScale(95, 105, { maxTicks: 6, includeZero: false })
    expect(scale.min).toBeGreaterThan(80)
    expect(scale.min).toBeLessThanOrEqual(95)
    expect(scale.max).toBeGreaterThanOrEqual(105)
  })

  it('keeps an author\u{2019}s axis range exactly, ticking the round values inside it', () => {
    const scale = niceScale(0, 240, { maxTicks: 5, min: 50, max: 300 })
    expect(scale.min).toBe(50)
    expect(scale.max).toBe(300)
    expect(scale.ticks.every((t) => t >= 50 && t <= 300)).toBe(true)
    expect(scale.ticks.every((t) => Math.abs(t / scale.step - Math.round(t / scale.step)) < 1e-9)).toBe(true)
  })

  it('handles a flat or empty range', () => {
    expect(niceScale(0, 0).max).toBeGreaterThan(0)
    expect(niceScale(5, 5).min).toBe(0)
    expect(niceScale(-5, -5).max).toBe(0)
  })

  it('has no floating-point noise in ticks', () => {
    const scale = niceScale(0, 0.7, { maxTicks: 8 })
    for (const t of scale.ticks) expect(String(t).length).toBeLessThan(6)
  })
})

describe('helpers', () => {
  it('counts the decimals a step needs', () => {
    expect(stepDecimals(50)).toBe(0)
    expect(stepDecimals(2.5)).toBe(1)
    expect(stepDecimals(0.25)).toBe(2)
  })

  it('projects values onto a coordinate range', () => {
    const scale = niceScale(0, 100, { min: 0, max: 100 })
    expect(project(scale, 25, 200, 0)).toBe(150)
    expect(project(scale, 100, 200, 0)).toBe(0)
  })
})
