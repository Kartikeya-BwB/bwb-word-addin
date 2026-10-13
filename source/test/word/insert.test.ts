import { afterEach, describe, expect, it } from 'vitest'
import { insertBlocks } from '../../src/word/insert.ts'
import { insertDisclaimer } from '../../src/word/parts.ts'
import { createModel, para } from '../fakeWord/model.ts'
import { installFakeWord, uninstallFakeWord } from '../fakeWord/install.ts'

afterEach(uninstallFakeWord)

const callout = [{ type: 'callout' as const, title: 'Key takeaway', text: 'Blended finance works when risk is priced.' }]

describe('insertBlocks', () => {
  it('replaces an empty paragraph at the cursor', async () => {
    const empty = para('')
    const model = createModel([para('Before'), empty, para('After')], { selection: [empty] })
    installFakeWord(model)
    await insertBlocks(callout)
    expect(model.inserts[0]?.location).toBe('Replace')
    expect(model.paragraphs[0]?.text).toBe('Before')
    expect(model.paragraphs.some((p) => p.text === 'Key takeaway')).toBe(true)
    expect(model.paragraphs.includes(empty)).toBe(false)
    expect(model.paragraphs[model.paragraphs.length - 1]?.text).toBe('After')
  })

  it('goes below a paragraph that has text, never into it', async () => {
    const first = para('A paragraph with text.')
    const model = createModel([first, para('Next')], { selection: [first] })
    installFakeWord(model)
    await insertBlocks(callout)
    expect(model.inserts[0]?.location).toBe('After')
    expect(model.paragraphs[0]).toBe(first)
    expect(first.text).toBe('A paragraph with text.')
  })

  it('goes after the table when the cursor is in a cell, so blocks are never nested', async () => {
    const table = { style: 'BwB Standard', rows: [['Country', 'Size']] }
    const cell = para('Country', 'Normal', { tableNestingLevel: 1, table })
    const model = createModel([para('Intro'), cell, para('After the table')], { selection: [cell] })
    model.tables.push(table)
    installFakeWord(model)
    await insertBlocks(callout)
    const texts = model.paragraphs.map((p) => p.text)
    expect(texts.indexOf('Key takeaway')).toBeGreaterThan(texts.indexOf('Country'))
    expect(model.paragraphs.filter((p) => p.text === 'Key takeaway').every((p) => p.tableNestingLevel <= 1)).toBe(true)
  })

  it('adds at the end of the document when asked', async () => {
    const model = createModel([para('Only')])
    installFakeWord(model)
    await insertBlocks(callout, 'end')
    expect(model.inserts[0]?.location).toBe('End')
  })

  it('numbers a new caption after the captions already in the document, then refreshes the SEQ fields', async () => {
    const model = createModel([
      para('Table 1. Earlier table', 'BwB Table Caption'),
      para('Table 2. Another', 'BwB Table Caption'),
      para(''),
    ])
    model.selection = model.paragraphs.slice(2)
    model.fields.push({ code: ' SEQ Table \\* ARABIC ', updates: 0 }, { code: ' TOC \\o "1-2" ', updates: 0 })
    installFakeWord(model)
    await insertBlocks([{ type: 'table', header: ['A', 'B'], rows: [['1', '2']], caption: 'New table' }])
    expect(model.inserts[0]?.ooxml).toMatch(/>3</)
    expect(model.fields.map((f) => f.updates)).toEqual([1, 0])
  })
})

describe('insertDisclaimer', () => {
  it('adds the disclaimer at the end once, and refuses a second copy', async () => {
    const model = createModel([para('Body text.')])
    installFakeWord(model)
    await insertDisclaimer()
    expect(model.inserts[0]?.location).toBe('End')
    expect(model.paragraphs.some((p) => p.text.includes('953637'))).toBe(true)
    await expect(insertDisclaimer()).rejects.toThrow(/already has the BwB disclaimer/)
  })
})
