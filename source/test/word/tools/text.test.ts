import { afterEach, describe, expect, it } from 'vitest'
import { clearStrayFormatting, joinBrokenLines, removeEmptyParagraphs } from '../../../src/word/tools/cleanup.ts'
import { addNonBreakingSpaces, changeSelectionCase, insertSymbol } from '../../../src/word/tools/textTools.ts'
import type { FakeParagraph, FakeTable } from '../../fakeWord/model.ts'
import { createModel, font, para } from '../../fakeWord/model.ts'
import { installFakeWord, uninstallFakeWord } from '../../fakeWord/install.ts'

afterEach(uninstallFakeWord)

/** The whole of one paragraph selected. */
const whole = (p: FakeParagraph): { paragraph: FakeParagraph; start: number; end: number } => ({ paragraph: p, start: 0, end: p.text.length })

describe('changeSelectionCase', () => {
  it('changes the selected words only, keeping acronyms', async () => {
    const p = para('We met The UNDP Team On Monday.')
    const model = createModel([p], { selectionRange: { paragraph: p, start: 7, end: 24 } })
    installFakeWord(model)
    expect(await changeSelectionCase('lower')).toBe('3 words changed to lower case. Check the names of people, places and organisations.')
    expect(p.text).toBe('We met the UNDP team on Monday.')
  })

  it('keeps the acronyms the document uses when the selection is all in capitals', async () => {
    const heading = para('UNICEF AND THE WATER SECTOR', 'heading 1')
    const model = createModel([heading, para('UNICEF works with national water utilities.')], { selectionRange: whole(heading) })
    installFakeWord(model)
    await changeSelectionCase('sentence')
    expect(heading.text).toBe('UNICEF and the water sector')
  })

  it('works paragraph by paragraph across a selection of several', async () => {
    const a = para('the role of the world bank')
    const b = para('what comes next')
    const model = createModel([a, b], { selection: [a, b], selectedText: 'the role of the world bank\rwhat comes next' })
    installFakeWord(model)
    expect(await changeSelectionCase('title')).toBe('7 words changed to title case.')
    expect([a.text, b.text]).toEqual(['The Role of the World Bank', 'What Comes Next'])
  })

  it('uses Word\u2019s own upper case on the desktop, and goes word by word elsewhere', async () => {
    const p = para('Blended finance')
    const model = createModel([p], { selectionRange: whole(p), desktopApiVersion: 1.4 })
    installFakeWord(model)
    expect(await changeSelectionCase('upper')).toBe('Selection changed to upper case.')
    expect(p.text).toBe('BLENDED FINANCE')
    p.text = 'Blended finance'
    model.desktopApiVersion = 0
    expect(await changeSelectionCase('upper')).toBe('2 words changed to upper case.')
    expect(p.text).toBe('BLENDED FINANCE')
  })

  it('needs a selection', async () => {
    installFakeWord(createModel([para('Text')]))
    await expect(changeSelectionCase('upper')).rejects.toThrow('Select the text to change first.')
  })
})

describe('addNonBreakingSpaces', () => {
  it('works on the whole document when nothing is selected', async () => {
    const a = para('It costs USD 5m, up 12 per cent since 5 October 2026.')
    const b = para('A 20 MW plant.')
    const c = para('Nothing here.')
    installFakeWord(createModel([a, b, c]))
    const message = await addNonBreakingSpaces()
    expect(message).toBe('Non-breaking spaces added in 4 places in the document, so figures, dates and units stay on one line.')
    expect(a.text).toBe('It costs USD\u00a05m, up 12\u00a0per\u00a0cent since 5\u00a0October\u00a02026.')
    expect(b.text).toBe('A 20\u00a0MW plant.')
  })

  it('works on the selected paragraphs only when there is a selection', async () => {
    const a = para('USD 5m here')
    const b = para('USD 6m there')
    installFakeWord(createModel([a, b], { selectionRange: { paragraph: a, start: 0, end: 4 } }))
    expect(await addNonBreakingSpaces()).toMatch(/in 1 place in the selection/)
    expect([a.text, b.text]).toEqual(['USD\u00a05m here', 'USD 6m there'])
  })
})

describe('insertSymbol', () => {
  it('inserts a symbol at the cursor', async () => {
    const p = para('Done ')
    installFakeWord(createModel([p], { selectionRange: { paragraph: p, start: 5, end: 5 } }))
    await insertSymbol('\u2713')
    expect(p.text).toBe('Done \u2713')
  })

  it('keeps the en dash for ranges in tables', async () => {
    const p = para('2020 2025')
    const table: FakeTable = { style: 'BwB Standard', rows: [['2020 2025']] }
    const model = createModel([p], { selectionRange: { paragraph: p, start: 4, end: 5 } })
    installFakeWord(model)
    await expect(insertSymbol('\u2013')).rejects.toThrow(/In running text, write the range with \u201cto\u201d/)
    model.selectionTable = table
    await insertSymbol('\u2013')
    expect(p.text).toBe('2020\u20132025')
  })
})

const LONG = 'Climate finance flows to sub-Saharan Africa grew in 2025 as the'

describe('joinBrokenLines', () => {
  it('rejoins lines pasted from a PDF, keeping real paragraph ends and collapsing double spaces', async () => {
    const lines = [para(LONG), para('Green Climate Fund scaled up its finan-'), para('cing for adaptation  projects.'), para('Next steps')]
    const model = createModel([para('Before'), ...lines, para('After')], { selection: lines })
    installFakeWord(model)
    expect(await joinBrokenLines()).toBe('1 broken line joined, 1 hyphenated word rejoined and 1 double space removed.')
    expect(model.paragraphs.map((p) => p.text)).toEqual([
      'Before',
      `${LONG} Green Climate Fund scaled up its financing for adaptation projects.`,
      'Next steps',
      'After',
    ])
  })

  it('joins manual line breaks inside a paragraph', async () => {
    const p = para(`${LONG}\u000bGreen Climate Fund grew.`)
    installFakeWord(createModel([p], { selection: [p] }))
    expect(await joinBrokenLines()).toBe('1 broken line joined.')
    expect(p.text).toBe(`${LONG} Green Climate Fund grew.`)
  })

  it('says when there is nothing to join, and refuses tables', async () => {
    const lines = [para('A full sentence.'), para('Another one.')]
    installFakeWord(createModel(lines, { selection: lines }))
    expect(await joinBrokenLines()).toBe('No broken lines in the selection.')
    const single = [para('One line')]
    installFakeWord(createModel(single, { selection: single }))
    await expect(joinBrokenLines()).rejects.toThrow('Select the lines to join first.')
  })
})

describe('clearStrayFormatting', () => {
  it('removes fonts, colours, sizes and highlighting, keeping bold', async () => {
    const p = para('Pasted text', 'Normal', { font: font({ name: 'Arial', color: '#FF0000', size: 14, bold: true, highlightColor: 'Yellow' }) })
    const model = createModel([p], { selectionRange: whole(p) })
    installFakeWord(model)
    expect(await clearStrayFormatting()).toBe('Stray formatting cleared from 1 paragraph. Bold, italic and underlining were kept.')
    expect(p.font).toEqual({ name: null, size: null, color: null, highlightColor: null, bold: true, italic: false })
    expect(p.text).toBe('Pasted text')
    expect(model.inserts[0]?.location).toBe('Inline')
  })

  it('leaves paragraphs with comments or tracked changes alone, and says when there is nothing to do', async () => {
    const commented = para('Has a comment', 'Normal', {
      font: font({ name: 'Arial' }),
      ooxml: '<w:p><w:commentRangeStart w:id="1"/><w:r><w:rPr><w:rFonts w:ascii="Arial"/></w:rPr><w:t>Has a comment</w:t></w:r></w:p>',
    })
    const plain = para('Already plain', 'Normal', { font: font({ name: null, size: null, color: null }) })
    const model = createModel([commented, plain], { selection: [commented, plain], selectedText: 'x' })
    installFakeWord(model)
    expect(await clearStrayFormatting()).toBe('No stray formatting to clear. 1 paragraph was left alone so its comments stay intact.')
    expect(commented.font.name).toBe('Arial')
  })
})

describe('removeEmptyParagraphs', () => {
  it('removes empty paragraphs, keeping those with a break, a picture or a section break, and between tables', async () => {
    const table: FakeTable = { style: 'BwB Standard', rows: [['a']] }
    const other: FakeTable = { style: 'BwB Standard', rows: [['b']] }
    const empty = para('')
    const pageBreak = para('', 'Normal', { ooxml: '<w:p><w:r><w:br w:type="page"/></w:r></w:p>' })
    const section = para('', 'Normal', { ooxml: '<w:p><w:pPr><w:sectPr/></w:pPr></w:p>' })
    const pictureLine = para('')
    const between = para('')
    const doc = [
      para('Start'), empty, pageBreak, section, pictureLine,
      para('a', 'Normal', { tableNestingLevel: 1, table }), between, para('b', 'Normal', { tableNestingLevel: 1, table: other }),
      para('  '), para('End'),
    ]
    const model = createModel(doc, { selection: doc })
    model.pictures.push({ altTextTitle: '', altTextDescription: 'Chart', paragraph: pictureLine })
    installFakeWord(model)
    expect(await removeEmptyParagraphs()).toBe(
      '2 empty paragraphs removed. 4 paragraphs that look empty but hold a picture, a break or a section break, or separate two tables, were kept.',
    )
    expect(model.paragraphs).not.toContain(empty)
    expect(model.paragraphs).toEqual(expect.arrayContaining([pageBreak, section, pictureLine, between]))
    expect(model.paragraphs.map((p) => p.text)).not.toContain('  ')
  })
})
