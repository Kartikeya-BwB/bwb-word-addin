import { afterEach, describe, expect, it } from 'vitest'
import type { FixAction } from '../../src/check/types.ts'
import { StaleDocumentError, applyFix, applyFixes, goTo } from '../../src/word/fixes.ts'
import { createModel, font, para } from '../fakeWord/model.ts'
import { installFakeWord, uninstallFakeWord } from '../fakeWord/install.ts'

afterEach(uninstallFakeWord)

function replace(paragraph: number, text: string, match: string, replacement: string, from = 0): FixAction {
  const start = text.indexOf(match, from)
  if (start < 0) throw new Error(`${match} not in ${text}`)
  return { kind: 'replace', paragraph, expectedText: text, start, end: start + match.length, match, replacement }
}

describe('applyFixes', () => {
  it('applies several replacements in one paragraph without the offsets colliding', async () => {
    const text = 'The organization will analyze the program, and the program will analyze it.'
    const p = para(text)
    installFakeWord(createModel([p]))
    const fixes = [
      replace(0, text, 'organization', 'organisation'),
      replace(0, text, 'analyze', 'analyse'),
      replace(0, text, 'program', 'programme'),
      replace(0, text, 'program', 'programme', text.indexOf('program') + 1),
      replace(0, text, 'analyze', 'analyse', text.indexOf('analyze') + 1),
    ]
    expect(await applyFixes(fixes)).toEqual({ applied: 5, skipped: 0 })
    expect(p.text).toBe('The organisation will analyse the programme, and the programme will analyse it.')
  })

  it('skips fixes for a paragraph that changed after the check, rather than guessing', async () => {
    const p = para('The color is wrong.')
    installFakeWord(createModel([p]))
    const stale = replace(0, 'The color was wrong.', 'color', 'colour')
    await expect(applyFix(stale)).rejects.toBeInstanceOf(StaleDocumentError)
    expect(p.text).toBe('The color is wrong.')
  })

  it('deletes paragraphs from the end of the document back, so the indexes stay right', async () => {
    const model = createModel([para('A'), para(''), para('B'), para('')])
    installFakeWord(model)
    const result = await applyFixes([
      { kind: 'delete-paragraph', paragraph: 1, expectedText: '' },
      { kind: 'delete-paragraph', paragraph: 3, expectedText: '' },
    ])
    expect(result).toEqual({ applied: 2, skipped: 0 })
    expect(model.paragraphs.map((p) => p.text)).toEqual(['A', 'B'])
  })

  it('changes a paragraph style, font and highlighting', async () => {
    const heading = para('Background')
    const stray = para('Stray Arial text', 'Normal', { font: font({ name: 'Arial', highlightColor: 'Yellow' }) })
    installFakeWord(createModel([heading, stray]))
    await applyFix({ kind: 'paragraph-style', paragraph: 0, expectedText: 'Background', style: 'heading 2', builtIn: 'Heading2' })
    await applyFix({ kind: 'paragraph-font', paragraph: 1, expectedText: 'Stray Arial text', font: 'Aptos' })
    await applyFix({ kind: 'remove-highlight', paragraph: 1, expectedText: 'Stray Arial text' })
    expect(heading.styleBuiltIn).toBe('Heading2')
    expect(stray.font.name).toBe('Aptos')
    expect(stray.font.highlightColor).toBeNull()
  })

  it('gives a table the BwB Standard style, bringing the styles in first', async () => {
    const model = createModel([para('x')])
    model.tables.push({ style: 'Table Grid', rows: [['a']] })
    installFakeWord(model)
    await applyFix({ kind: 'table-style', table: 0, style: 'BwB Standard' })
    expect(model.tables[0]?.style).toBe('BwB Standard')
  })

  it('removes the writing prompts and their text', async () => {
    const prompt = para('Say what the report is for.', 'BwB Lead')
    const model = createModel([para('Intro'), prompt, para('Body')])
    model.controls.push({ tag: 'BwBGuidance', title: 'Writing prompt', paragraphs: [prompt] })
    installFakeWord(model)
    await applyFix({ kind: 'remove-controls', tag: 'BwBGuidance' })
    expect(model.paragraphs.map((p) => p.text)).toEqual(['Intro', 'Body'])
    expect(model.controls).toHaveLength(0)
  })

  it('fills in document properties', async () => {
    const model = createModel()
    installFakeWord(model)
    await applyFix({ kind: 'properties', values: { author: 'Bankers without Boundaries', company: 'Bankers without Boundaries' } })
    expect(model.properties).toMatchObject({ author: 'Bankers without Boundaries', company: 'Bankers without Boundaries' })
  })
})

describe('goTo', () => {
  it('selects the matched words when they can be found, or else the paragraph', async () => {
    const text = 'We analyze data.'
    const model = createModel([para(text)])
    installFakeWord(model)
    await goTo({ kind: 'paragraph', paragraph: 0, expectedText: text, start: 3, end: 10 })
    await goTo({ kind: 'paragraph', paragraph: 0, expectedText: text })
    expect(model.selected).toEqual(['range:Select:analyze', `paragraph:${text}`])
  })

  it('refuses to go to a paragraph that has changed', async () => {
    installFakeWord(createModel([para('New text')]))
    await expect(goTo({ kind: 'paragraph', paragraph: 0, expectedText: 'Old text' })).rejects.toBeInstanceOf(StaleDocumentError)
  })

  it('selects tables by position', async () => {
    const model = createModel([para('x')])
    model.tables.push({ style: 'BwB Standard', rows: [['a']] })
    installFakeWord(model)
    await goTo({ kind: 'table', table: 0 })
    expect(model.selected).toEqual(['table:BwB Standard'])
  })
})
