import { describe, expect, it } from 'vitest'
import { chooseUnit, dayNumber, formatDay, isoDay, londonToday, parseDate, periods } from '../../src/charts/dates.ts'
import { ganttRows, milestoneLabel } from '../../src/charts/types/gantt.ts'

const d = (y: number, m: number, day: number): number => dayNumber(y, m - 1, day)

describe('parseDate', () => {
  it.each([
    ['5 October 2026', d(2026, 10, 5)],
    ['5 Oct 2026', d(2026, 10, 5)],
    ['5th October 2026', d(2026, 10, 5)],
    ['05-Oct-26', d(2026, 10, 5)],
    ['2026-10-05', d(2026, 10, 5)],
    ['2026/10/05', d(2026, 10, 5)],
    ['05/10/2026', d(2026, 10, 5)],
    ['5.10.2026', d(2026, 10, 5)],
    ['10/25/2026', d(2026, 10, 25)],
    ['October 5, 2026', d(2026, 10, 5)],
    ['46300', d(2026, 10, 5)],
  ])('reads %s as a day', (text, day) => {
    expect(parseDate(text)).toEqual({ first: day, last: day, precision: 'day' })
  })

  it('reads months, quarters, halves and years as spans', () => {
    expect(parseDate('October 2026')).toEqual({ first: d(2026, 10, 1), last: d(2026, 10, 31), precision: 'month' })
    expect(parseDate('Feb 2028')).toEqual({ first: d(2028, 2, 1), last: d(2028, 2, 29), precision: 'month' })
    expect(parseDate('Oct-26')).toMatchObject({ first: d(2026, 10, 1), precision: 'month' })
    expect(parseDate('2026-10')).toMatchObject({ first: d(2026, 10, 1), precision: 'month' })
    expect(parseDate('Q4 2026')).toEqual({ first: d(2026, 10, 1), last: d(2026, 12, 31), precision: 'quarter' })
    expect(parseDate('2027 Q1')).toEqual({ first: d(2027, 1, 1), last: d(2027, 3, 31), precision: 'quarter' })
    expect(parseDate('H2 2027')).toEqual({ first: d(2027, 7, 1), last: d(2027, 12, 31), precision: 'half' })
    expect(parseDate('2027')).toEqual({ first: d(2027, 1, 1), last: d(2027, 12, 31), precision: 'year' })
  })

  it('rejects text that is not a date', () => {
    for (const text of ['', 'soon', '31 February 2026', '2026-13-01', 'Mar', '32/01/2026', 'Smarch 2026']) expect(parseDate(text)).toBeNull()
  })
})

describe('time axis', () => {
  it('lists the months, quarters or years a plan covers', () => {
    const months = periods(d(2026, 11, 15), d(2027, 2, 3), 'month')
    expect(months.map((p) => p.labels[0])).toEqual(['Nov', 'Dec', 'Jan', 'Feb'])
    expect(months.map((p) => p.year)).toEqual([2026, 2026, 2027, 2027])
    expect(months[0]?.start).toBe(d(2026, 11, 1))
    expect(months[3]?.end).toBe(d(2027, 3, 1))
    const quarters = periods(d(2026, 11, 15), d(2027, 8, 1), 'quarter')
    expect(quarters.map((p) => `${p.labels[0]} ${p.year}`)).toEqual(['Q4 2026', 'Q1 2027', 'Q2 2027', 'Q3 2027'])
    expect(periods(d(2026, 5, 1), d(2028, 1, 1), 'year').map((p) => p.labels[0])).toEqual(['2026', '2027', '2028'])
  })

  it('chooses months for short plans, quarters for a few years, then years', () => {
    expect(chooseUnit(d(2026, 11, 1), d(2027, 4, 30))).toBe('month')
    expect(chooseUnit(d(2026, 1, 1), d(2028, 6, 30))).toBe('quarter')
    expect(chooseUnit(d(2026, 1, 1), d(2032, 12, 31))).toBe('year')
  })

  it('writes days the BwB way', () => {
    expect(formatDay(d(2026, 10, 5))).toBe('5 October 2026')
    expect(formatDay(d(2026, 10, 5), 'short')).toBe('5 Oct')
    expect(formatDay(d(2026, 10, 5), 'month')).toBe('October 2026')
    expect(isoDay(d(2026, 1, 9))).toBe('2026-01-09')
  })

  it('takes today in London, where midnight UTC in summer is already the next day', () => {
    expect(londonToday(new Date('2026-06-30T23:30:00Z'))).toBe(d(2026, 7, 1))
    expect(londonToday(new Date('2026-12-31T23:30:00Z'))).toBe(d(2026, 12, 31))
  })
})

describe('ganttRows', () => {
  it('reads tasks, milestones and group headings', () => {
    const rows = ganttRows([
      ['Task', 'Start', 'End'],
      ['Inception', '', ''],
      ['Review', '2 November 2026', 'November 2026'],
      ['Report', '30 November 2026', ''],
      ['Close', '', 'Q1 2027'],
    ])
    expect(rows).toEqual([
      { label: 'Inception', kind: 'group', start: 0, end: 0, indent: false },
      { label: 'Review', kind: 'task', start: d(2026, 11, 2), end: d(2026, 11, 30), indent: true },
      { label: 'Report', kind: 'milestone', start: d(2026, 11, 30), end: d(2026, 11, 30), indent: true, precision: 'day' },
      { label: 'Close', kind: 'milestone', start: d(2027, 3, 31), end: d(2027, 3, 31), indent: true, precision: 'quarter' },
    ])
  })

  it('labels milestones as precisely as their dates were given', () => {
    const rows = ganttRows([['Task', 'Date'], ['A', '30 November 2026'], ['B', 'March 2027'], ['C', 'Q2 2027'], ['D', 'H2 2027'], ['E', '2028']])
    expect(rows.map((r) => milestoneLabel(r, false))).toEqual(['30 Nov', 'Mar 2027', 'Q2 2027', 'H2 2027', '2028'])
    const first = rows[0]
    if (first) expect(milestoneLabel(first, true)).toBe('30 Nov 2026')
    expect(rows[1]?.start).toBe(d(2027, 3, 31))
  })

  it('swaps dates typed the wrong way round and skips rows with bad dates', () => {
    const warnings: string[] = []
    const rows = ganttRows([['Task', 'Start', 'End'], ['Late', '2027-03-01', '2027-01-01'], ['Odd', 'soon', '2027-01-01']], (m) => warnings.push(m))
    expect(rows).toEqual([{ label: 'Late', kind: 'task', start: d(2027, 1, 1), end: d(2027, 3, 1), indent: false }])
    expect(warnings[0]).toContain('soon')
  })
})
