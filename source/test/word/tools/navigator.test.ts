import { afterEach, describe, expect, it } from 'vitest'
import type { OutlineItem } from '../../../src/word/tools/navigator.ts'
import {
  bookmarkBase, goToItem, insertCrossReference, readCursorHeading, readOutline, shiftHeadingLevel,
} from '../../../src/word/tools/navigator.ts'
import type { FakeModel, FakeParagraph, FakeTable } from '../../fakeWord/model.ts'
import { createModel, para } from '../../fakeWord/model.ts'
import { installFakeWord, uninstallFakeWord } from '../../fakeWord/install.ts'

afterEach(uninstallFakeWord)

const cells = (table: FakeTable): FakeParagraph[] => table.rows.flatMap((r) => r.map((t) => para(t, 'Normal', { tableNestingLevel: 1, table })))

function report(): { model: FakeModel; heading: FakeParagraph; cursor: FakeParagraph } {
  const data: FakeTable = { style: 'BwB Standard', rows: [['Country', 'Size'], ['Kenya', '250']] }
  const callout: FakeTable = { style: 'BwB Highlight', rows: [['Key takeaway']] }
  const bare: FakeTable = { style: 'Table Grid', rows: [['Risk', 'Rating']] }
  const heading = para('Background', 'heading 2')
  const cursor = para('Body text where a reference goes.')
  const model = createModel([
    para('Introduction', 'heading 1'),
    cursor,
    heading,
    ...cells(data),
    para('Table 1. Green bond issuance by country', 'BwB Table Caption'),
    ...cells(callout),
    para(''),
    para('Figure 1. Climate finance flows, 2025', 'BwB Image Caption'),
    para('Annex', 'BwB Heading 1 Unnumbered'),
    ...cells(bare),
  ])
  model.tables.push(data, callout, bare)
  model.styles.add('BwB Heading 1 Unnumbered').add('BwB Heading 2 Unnumbered')
  return { model, heading, cursor }
}

describe('readOutline', () => {
  it('lists headings, data tables and figures in order, nesting tables and figures under their heading', async () => {
    installFakeWord(report().model)
    const outline = await readOutline()
    expect(outline.map((i) => [i.kind, i.level, i.label])).toEqual([
      ['heading', 1, 'Introduction'],
      ['heading', 2, 'Background'],
      ['table', 3, 'Table 1. Green bond issuance by country'],
      ['figure', 3, 'Figure 1. Climate finance flows, 2025'],
      ['heading', 1, 'Annex'],
      ['table', 2, 'Table with no caption: Risk'],
    ])
    expect(outline[2]).toMatchObject({ table: 0, paragraph: 7, number: 'Table 1' })
    expect(outline[3]).toMatchObject({ number: 'Figure 1' })
    expect(outline[5]).toMatchObject({ table: 2, paragraph: -1 })
  })
})

describe('goToItem', () => {
  it('goes to a heading or a table, and refuses an item the document no longer matches', async () => {
    const { model } = report()
    installFakeWord(model)
    const [intro, , table] = await readOutline()
    if (!intro || !table) throw new Error('outline too short')
    await goToItem(intro)
    await goToItem(table)
    expect(model.selected).toEqual(['paragraph:Introduction', 'table:BwB Standard'])
    expect(model.selection.map((p) => p.text)).toEqual(['Introduction'])
    const stale: OutlineItem = { ...intro, expectedText: 'Old heading' }
    await expect(goToItem(stale)).rejects.toThrow('The document has changed. Refresh the navigator.')
  })
})

describe('shiftHeadingLevel', () => {
  it('promotes and demotes the heading at the cursor, within the BwB levels', async () => {
    const { model, heading } = report()
    model.selection = [heading]
    installFakeWord(model)
    expect(await readCursorHeading()).toEqual({ level: 2, text: 'Background' })
    expect(await shiftHeadingLevel(-1)).toBe('Heading 2 is now Heading 1.')
    expect(heading.styleBuiltIn).toBe('Heading1')
    await expect(shiftHeadingLevel(-1)).rejects.toThrow('Heading 1 is the top level already.')
    expect(await shiftHeadingLevel(1)).toBe('Heading 1 is now Heading 2.')
    expect(heading.style).toBe('heading 2')
  })

  it('moves unnumbered BwB headings between their own styles, and needs a heading', async () => {
    const { model, cursor } = report()
    const annex = model.paragraphs.find((p) => p.text === 'Annex')
    if (!annex) throw new Error('no annex')
    model.selection = [annex]
    installFakeWord(model)
    await shiftHeadingLevel(1)
    expect(annex.style).toBe('BwB Heading 2 Unnumbered')
    model.selection = [cursor]
    expect(await readCursorHeading()).toBeNull()
    await expect(shiftHeadingLevel(1)).rejects.toThrow('Put the cursor in a heading first.')
  })
})

describe('insertCrossReference', () => {
  it('bookmarks a heading once and inserts a REF field to it at the cursor', async () => {
    const { model, cursor, heading } = report()
    model.selection = [cursor]
    installFakeWord(model)
    const item = (await readOutline())[1]
    if (!item) throw new Error('no heading')
    expect(await insertCrossReference(item)).toBe('Cross-reference to \u201cBackground\u201d inserted. It updates with the document\u2019s fields (F9).')
    expect(model.bookmarks?.get('BwB_Background')).toEqual({ paragraph: heading, start: 0, end: 10 })
    expect(model.fields.map((f) => [f.code, f.updates])).toEqual([[' REF BwB_Background \\h ', 1]])
    await insertCrossReference(item, 'number')
    expect([...(model.bookmarks?.keys() ?? [])]).toEqual(['BwB_Background'])
    expect(model.fields[1]?.code).toBe(' REF BwB_Background \\r \\h ')
  })

  it('refers to a caption by its label, choosing a bookmark name not taken', async () => {
    const { model } = report()
    const caption = model.paragraphs.find((p) => p.text.startsWith('Table 1.'))
    if (!caption) throw new Error('no caption')
    model.bookmarks = new Map([['BwB_Table_1', { paragraph: para('Elsewhere'), start: 0, end: 9 }]])
    installFakeWord(model)
    const table = (await readOutline()).find((i) => i.kind === 'table')
    if (!table) throw new Error('no table')
    expect(await insertCrossReference(table)).toMatch(/^Cross-reference to Table 1 inserted\./)
    expect(model.bookmarks.get('BwB_Table_1_2')).toEqual({ paragraph: caption, start: 0, end: 7 })
  })

  it('explains that Word on the web cannot insert cross-references, and needs a caption to refer to', async () => {
    const { model } = report()
    installFakeWord(model)
    const outline = await readOutline()
    const bare = outline[outline.length - 1]
    const heading = outline[0]
    if (!bare || !heading) throw new Error('outline too short')
    await expect(insertCrossReference(bare)).rejects.toThrow('This table has no caption to refer to. Add a caption below it first.')
    model.platform = 'OfficeOnline'
    model.desktopApiVersion = 0
    await expect(insertCrossReference(heading)).rejects.toThrow('Cross-references need Word for Windows or Mac: Word on the web cannot add fields.')
  })

  it('makes bookmark names Word accepts', () => {
    expect(bookmarkBase('Table 3')).toBe('BwB_Table_3')
    expect(bookmarkBase('C\u00f4te d\u2019Ivoire: the 2025 outlook for blended finance')).toBe('BwB_C_te_d_Ivoire_the_2025_outlook_f')
    expect(bookmarkBase('')).toBe('BwB')
  })
})
