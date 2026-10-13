import { describe, expect, it } from 'vitest'
import { STEP_COLOURS, stepColour, waterfallSteps } from '../../src/charts/types/waterfall.ts'

const bridge = [
  ['Step', 'Value', 'Kind'],
  ['Budget 2025', '100', 'total'],
  ['Staff', '12', ''],
  ['Travel', '-5', ''],
  ['Before savings', '', 'subtotal'],
  ['Efficiencies', '-6', 'delta'],
  ['Budget 2026', '', 'total'],
]

describe('waterfallSteps', () => {
  it('runs the cumulative total from the start through each change', () => {
    const steps = waterfallSteps(bridge)
    expect(steps.map((s) => [s.label, s.kind, s.from, s.to])).toEqual([
      ['Budget 2025', 'total', 0, 100],
      ['Staff', 'delta', 100, 112],
      ['Travel', 'delta', 112, 107],
      ['Before savings', 'subtotal', 0, 107],
      ['Efficiencies', 'delta', 107, 101],
      ['Budget 2026', 'total', 0, 101],
    ])
  })

  it('uses a typed total as the new running total', () => {
    const steps = waterfallSteps([['Step', 'Value', 'Kind'], ['Start', '50', 't'], ['Up', '10', ''], ['Restated', '70', 'total'], ['Down', '-20', '']])
    expect(steps[3]).toMatchObject({ from: 70, to: 50 })
  })

  it('without a kind column, takes the first row as the start and empty values as totals', () => {
    const steps = waterfallSteps([['Step', 'Value'], ['2024', '40'], ['Price', '5'], ['Volume', '-3'], ['2025', '']])
    expect(steps.map((s) => s.kind)).toEqual(['total', 'delta', 'delta', 'total'])
    expect(steps[3]?.to).toBe(42)
  })

  it('crosses zero', () => {
    const steps = waterfallSteps([['Step', 'Value', 'Kind'], ['Start', '10', 'total'], ['Loss', '-25', ''], ['End', '', 'total']])
    expect(steps[1]).toMatchObject({ from: 10, to: -15 })
    expect(steps[2]).toMatchObject({ from: 0, to: -15 })
  })

  it('reads units in the values and warns about text that is not a number', () => {
    const warnings: string[] = []
    const steps = waterfallSteps([['Step', 'Value'], ['Start', 'USD 1.5m'], ['Odd', 'lots'], ['Up', '(0.5m)']], (m) => warnings.push(m))
    expect(steps.map((s) => s.to)).toEqual([1.5, 1])
    expect(warnings).toHaveLength(1)
  })

  it('colours increases green, decreases charcoal and totals navy', () => {
    const steps = waterfallSteps(bridge)
    expect(steps.map(stepColour)).toEqual([
      STEP_COLOURS.total, STEP_COLOURS.increase, STEP_COLOURS.decrease, STEP_COLOURS.total, STEP_COLOURS.decrease, STEP_COLOURS.total,
    ])
    expect(STEP_COLOURS).toEqual({ increase: '328840', decrease: '4A4A4C', total: '022338' })
  })
})
