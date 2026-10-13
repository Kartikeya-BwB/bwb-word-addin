import { describe, expect, it } from 'vitest'
import { abbreviateMonths } from '../../src/charts/dates.ts'
import { ellipsise, fitLabels, textWidth, wrapText } from '../../src/charts/measure.ts'

describe('text measurement', () => {
  it('estimates widths in points from the font size', () => {
    expect(textWidth('', 9)).toBe(0)
    expect(textWidth('0000', 10)).toBeCloseTo(22.4, 5)
    expect(textWidth('WWW', 9)).toBeGreaterThan(textWidth('iii', 9))
    expect(textWidth('Total', 9, true)).toBeGreaterThan(textWidth('Total', 9))
    expect(textWidth('Total', 18)).toBeCloseTo(textWidth('Total', 9) * 2, 5)
  })

  it('wraps at spaces and after hyphens, never wider than allowed', () => {
    const wrapped = wrapText('Sub-Saharan Africa and South Asia', 60, 9)
    expect(wrapped.lines.length).toBeGreaterThan(1)
    for (const line of wrapped.lines) expect(textWidth(line, 9)).toBeLessThanOrEqual(60)
    expect(wrapped.truncated).toBe(false)
    expect(wrapText('Sub-Saharan', 45, 9).lines).toEqual(['Sub-', 'Saharan'])
  })

  it('ends with an ellipsis when the text needs more lines than allowed', () => {
    const wrapped = wrapText('one two three four five six seven eight nine ten', 40, 9, 2)
    expect(wrapped.lines).toHaveLength(2)
    expect(wrapped.lines[1]?.endsWith('\u{2026}')).toBe(true)
    expect(wrapped.truncated).toBe(true)
    expect(ellipsise('Netherlands', 30, 9).endsWith('\u{2026}')).toBe(true)
    expect(ellipsise('UK', 30, 9)).toBe('UK')
  })

  it('shrinks a set of labels together, down to 8 pt, before cutting any short', () => {
    const fitted = fitLabels(['Renewable energy', 'Water'], 72, 12, 1)
    expect(fitted.size).toBeLessThan(12)
    expect(fitted.size).toBeGreaterThanOrEqual(8)
    expect(fitted.wrapped.every((w) => !w.truncated)).toBe(true)
    const floor = fitLabels(['Extraordinarily long single word'], 20, 12, 1)
    expect(floor.size).toBe(8)
    expect(floor.wrapped[0]?.truncated).toBe(true)
  })

  it('shortens month names only when every label is one', () => {
    expect(abbreviateMonths(['January', 'February 2027', 'september'])).toEqual(['Jan', 'Feb 2027', 'Sep'])
    expect(abbreviateMonths(['January', 'Budget'])).toBeNull()
    expect(abbreviateMonths(['May', 'June'])).toEqual(['May', 'Jun'])
  })
})
