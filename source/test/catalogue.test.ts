import { describe, expect, it } from 'vitest'
import { searchTools, toolsFor } from '../src/commands/catalogue.ts'

describe('the tool catalogue', () => {
  it('lists every ribbon tool of each app with where it lives', () => {
    const word = toolsFor('word')
    const ppt = toolsFor('powerpoint')
    expect(word.length).toBeGreaterThan(60)
    expect(ppt.length).toBeGreaterThan(60)
    expect(word.find((t) => t.fn === 'styleHeading1')?.path).toBe('Text, Headings')
    expect(ppt.find((t) => t.fn === 'pptAlignLeft')?.path).toBe('Align, Align')
    expect(ppt.every((t) => Boolean(t.fn) !== Boolean(t.pane))).toBe(true)
  })

  it('finds tools by what they do, best match first', () => {
    const ppt = toolsFor('powerpoint')
    expect(searchTools(ppt, 'swap')[0]?.fn).toBe('pptSwap')
    expect(searchTools(ppt, 'harvey half')[0]?.fn).toBe('pptHarvey2')
    expect(searchTools(ppt, 'distribute').map((t) => t.fn)).toContain('pptDistributeHorizontally')
    const word = toolsFor('word')
    expect(searchTools(word, 'total row')[0]?.fn).toBe('wordTableAddTotal')
    expect(searchTools(word, 'pdf')[0]?.fn).toBe('wordJoinBrokenLines')
  })

  it('needs every word of the query to match somewhere, and caps the list', () => {
    expect(searchTools(toolsFor('word'), 'heading zebra')).toEqual([])
    expect(searchTools(toolsFor('powerpoint'), 'a').length).toBeLessThanOrEqual(8)
    expect(searchTools(toolsFor('word'), '   ')).toEqual([])
  })
})
