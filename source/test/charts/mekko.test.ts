import { describe, expect, it } from 'vitest'
import { createContext } from '../../src/charts/spec.ts'
import { prepareSeries } from '../../src/charts/series.ts'
import { mekkoColumns } from '../../src/charts/types/mekko.ts'

function columns(data: string[][]) {
  const ctx = createContext({ type: 'mekko', data, size: { widthPt: 470, heightPt: 255 } })
  return mekkoColumns(prepareSeries(ctx, { sortable: true, highlight: false }))
}

describe('mekkoColumns', () => {
  const data = [['Region', 'Energy', 'Water'], ['North', '60', '40'], ['South', '30', '20'], ['East', '10', '40']]

  it('makes column widths proportional to category totals', () => {
    const result = columns(data)
    expect(result.map((c) => c.total)).toEqual([100, 50, 50])
    expect(result.map((c) => c.width)).toEqual([0.5, 0.25, 0.25])
    expect(result.map((c) => c.start)).toEqual([0, 0.5, 0.75])
  })

  it('stacks each column to 100%', () => {
    for (const column of columns(data)) {
      const shares = column.segments.map((s) => s.share)
      expect(shares.reduce((a, b) => a + b, 0)).toBeCloseTo(1, 12)
      expect(column.segments[column.segments.length - 1]?.to).toBeCloseTo(1, 12)
    }
    expect(columns(data)[2]?.segments.map((s) => s.share)).toEqual([0.2, 0.8])
  })

  it('counts negative and missing values as zero', () => {
    const result = columns([['Region', 'A', 'B'], ['North', '-5', '10'], ['South', '', '10']])
    expect(result.map((c) => c.width)).toEqual([0.5, 0.5])
    expect(result[0]?.segments).toHaveLength(1)
  })
})
