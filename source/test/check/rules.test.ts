import { describe, expect, it } from 'vitest'
import { brandFindings } from '../../src/check/rulesBrand.ts'
import { consistencyFindings } from '../../src/check/rulesConsistency.ts'
import { sendFindings } from '../../src/check/rulesSend.ts'
import { structureFindings } from '../../src/check/rulesStructure.ts'
import { tableFindings } from '../../src/check/rulesTables.ts'
import { writingFindings } from '../../src/check/rulesWriting.ts'
import type { DocSnapshot, Finding, ParagraphInfo, TableInfo } from '../../src/check/types.ts'

const BWB = 'Bankers without Boundaries'
type P = Partial<ParagraphInfo> & { text: string }

function snap(paragraphs: P[], over: Partial<DocSnapshot> = {}): DocSnapshot {
  return {
    paragraphs: paragraphs.map((p, index) => ({
      index, style: 'Normal', styleBuiltIn: 'Normal', isListItem: false, inTable: false,
      font: { name: 'Aptos', size: 10, color: '#000000', highlight: null, bold: false }, ...p,
    })),
    tables: [], pictures: [], controls: [],
    headersFooters: [{ section: 0, kind: 'footer', text: 'Page 1 of 3', pictures: 0 }],
    fieldCodes: [], comments: 0, trackedChanges: 0, formatting: null,
    info: { title: 'A title', subject: '', author: BWB, company: BWB, docType: '', project: '', deliverable: '', version: '', funder: '', clientTemplate: false },
    fileName: null, bwbStyles: true, ...over,
  }
}

const h = (level: number, text: string, over: Partial<ParagraphInfo> = {}): P => ({ text, style: `heading ${level}`, styleBuiltIn: `Heading${level}`, ...over })
const rules = (findings: Finding[]): string[] => findings.map((f) => f.rule)
const table = (over: Partial<TableInfo>): TableInfo => ({ index: 0, style: 'BwB Standard', rows: 3, firstCell: 'Country', after: null, before: null, ...over })

describe('structure', () => {
  it('flags a skipped heading level', () => {
    const f = structureFindings(snap([h(1, 'Introduction'), h(3, 'Detail')]))
    expect(f.find((x) => x.rule === 'structure/skipped-level')?.title).toBe('Heading 3 follows heading 1')
  })

  it('flags typed heading numbers but not headings that start with a year, and only offers a fix when Word numbers the heading', () => {
    const typed = structureFindings(snap([h(1, '1. Introduction'), h(1, '2030 targets'), h(1, '2. Method', { isListItem: true })]))
      .filter((x) => x.rule === 'structure/typed-number')
    expect(typed.map((x) => x.excerpt)).toEqual(['1. Introduction', '2. Method'])
    expect(typed[0]?.fix).toBeUndefined()
    expect(typed[1]?.fix).toMatchObject({ kind: 'replace', start: 0, end: 3, replacement: '' })
  })

  it('removes a full stop at the end of a heading', () => {
    const f = structureFindings(snap([h(2, 'Scope.')])).find((x) => x.rule === 'structure/heading-full-stop')
    expect(f?.fix).toMatchObject({ kind: 'replace', start: 5, end: 6, replacement: '' })
  })

  it('suggests sentence case for a heading in title case, as advice', () => {
    const f = structureFindings(snap([h(2, 'Climate Finance For Small Island States')])).find((x) => x.rule === 'structure/heading-case')
    expect(f?.advisory).toBe(true)
    expect(f?.fix).toMatchObject({ kind: 'replace', replacement: 'Climate finance for small island states' })
  })

  it('flags the second empty paragraph in a row, not a single one', () => {
    expect(rules(structureFindings(snap([{ text: 'a' }, { text: '' }, { text: 'b' }])))).not.toContain('structure/empty-paragraph')
    const f = structureFindings(snap([{ text: 'a' }, { text: '' }, { text: '' }])).filter((x) => x.rule === 'structure/empty-paragraph')
    expect(f).toHaveLength(1)
    expect(f[0]?.fix).toEqual({ kind: 'delete-paragraph', paragraph: 2, expectedText: '' })
  })

  it('spots bold body text used as a heading, and very long paragraphs', () => {
    const long = Array.from({ length: 210 }, () => 'word').join(' ')
    const f = structureFindings(snap([{ text: 'Key findings', font: { name: 'Aptos', size: 10, color: null, highlight: null, bold: true } }, { text: 'Text.' }, { text: long }]))
    expect(rules(f)).toEqual(expect.arrayContaining(['structure/fake-heading', 'structure/long-paragraph']))
  })

  it('asks for contents in a report with five or more main headings, unless it has a TOC field', () => {
    const ps = [{ text: 'Report', styleBuiltIn: 'Title', style: 'Title' }, ...['A', 'B', 'C', 'D', 'E'].map((t) => h(1, t))]
    expect(rules(structureFindings(snap(ps)))).toContain('structure/no-contents')
    expect(rules(structureFindings(snap(ps, { fieldCodes: [' TOC \\o "1-2" \\h '] })))).not.toContain('structure/no-contents')
  })

  it('needs the disclaimer in client-facing types only, and calls it major there', () => {
    const final = structureFindings(snap([{ text: 'Body' }], { info: { ...snap([]).info, docType: 'final-report' } }))
    expect(final.find((x) => x.rule === 'structure/no-disclaimer')).toMatchObject({ severity: 'major', group: 'send', fix: { kind: 'insert-disclaimer' } })
    const eoi = structureFindings(snap([{ text: 'Body' }], { info: { ...snap([]).info, docType: 'eoi' } }))
    expect(rules(eoi)).not.toContain('structure/no-disclaimer')
    const withIt = structureFindings(snap([{ text: 'BwB (FRN: 953637) is authorised and regulated by the UK Financial Conduct Authority.' }], { info: { ...snap([]).info, docType: 'final-report' } }))
    expect(rules(withIt)).not.toContain('structure/no-disclaimer')
  })

  it('asks for the EU acknowledgement in EU grant deliverables', () => {
    const info = { ...snap([]).info, funder: 'eu-grant' as const }
    expect(rules(structureFindings(snap([{ text: 'Body' }], { info })))).toContain('structure/eu-acknowledgement')
    expect(rules(structureFindings(snap([{ text: 'Funded by the European Union. Views and opinions...' }], { info })))).not.toContain('structure/eu-acknowledgement')
  })

  it('flags front matter only when Word actually numbers it', () => {
    expect(rules(structureFindings(snap([h(1, 'Executive summary')])))).not.toContain('structure/numbered-front-matter')
    const f = structureFindings(snap([h(1, 'Executive summary', { isListItem: true })])).find((x) => x.rule === 'structure/numbered-front-matter')
    expect(f?.fix).toMatchObject({ kind: 'paragraph-style', style: 'BwB Heading 1 Unnumbered' })
  })
})

describe('tables and pictures', () => {
  it('offers BwB Standard for a table in another style, and asks for a caption below', () => {
    const f = tableFindings(snap([], { tables: [table({ style: 'Table Grid' })] }))
    expect(f.find((x) => x.rule === 'tables/style')?.fix).toEqual({ kind: 'table-style', table: 0, style: 'BwB Standard' })
    expect(rules(f)).toContain('tables/no-caption')
  })

  it('accepts a caption below, notes one above, and leaves abbreviation and layout tables alone', () => {
    const below = { text: 'Table 1. Flows by region. Source: OECD (2026)', style: 'BwB Table Caption' }
    expect(tableFindings(snap([], { tables: [table({ after: below })] }))).toEqual([])
    const above = tableFindings(snap([], { tables: [table({ before: below })] }))
    expect(above.map((x) => [x.rule, x.severity])).toEqual([['tables/caption-above', 'copyedit']])
    expect(tableFindings(snap([], { tables: [table({ firstCell: 'Abbreviation' })] }))).toEqual([])
    expect(tableFindings(snap([], { tables: [table({ style: 'BwB Highlight' })] }))).toEqual([])
    expect(tableFindings(snap([], { tables: [table({ rows: 1 })] }))).toEqual([])
  })

  it('asks for alt text on pictures without it', () => {
    const f = tableFindings(snap([], { pictures: [{ index: 0, altTitle: '', altDescription: '' }, { index: 1, altTitle: 'Map', altDescription: '' }] }))
    expect(f.map((x) => x.location)).toEqual([{ kind: 'picture', picture: 0 }])
  })
})

describe('brand', () => {
  it('notices a document not built on the template, unless a client template applies or Word cannot tell', () => {
    expect(brandFindings(snap([{ text: 'x' }], { bwbStyles: false })).find((x) => x.rule === 'brand/not-template')?.fix).toEqual({ kind: 'add-styles' })
    const client = snap([{ text: 'x' }], { bwbStyles: false })
    client.info.clientTemplate = true
    expect(rules(brandFindings(client))).not.toContain('brand/not-template')
    expect(rules(brandFindings(snap([{ text: 'x' }], { bwbStyles: null })))).not.toContain('brand/not-template')
  })

  it('fixes stray fonts paragraph by paragraph, choosing Candara for headings', () => {
    const arial = { name: 'Arial', size: 10, color: null, highlight: null, bold: false }
    const f = brandFindings(snap([{ text: 'Body', font: arial }, { ...h(1, 'Heading'), font: arial }]))
    expect(f.filter((x) => x.rule === 'brand/font').map((x) => x.fix)).toEqual([
      { kind: 'paragraph-font', paragraph: 0, expectedText: 'Body', font: 'Aptos' },
      { kind: 'paragraph-font', paragraph: 1, expectedText: 'Heading', font: 'Candara' },
    ])
  })

  it('treats a document mostly in another font as one style problem', () => {
    const ps = Array.from({ length: 30 }, (_, i) => ({ text: `Paragraph ${i}`, font: { name: 'Calibri', size: 11, color: null, highlight: null, bold: false } }))
    expect(rules(brandFindings(snap(ps)))).toEqual(['brand/document-font'])
  })

  it('checks colours, highlighting and text size from the run-level scan', () => {
    const f = brandFindings(snap([{ text: 'Coloured' }, { text: 'In a table', inTable: true }], {
      formatting: [
        { paragraph: 0, fonts: [], colours: ['FF0000', 'BED62F', '022338'], highlights: ['yellow'], smallestHalfPoints: 12, pictures: 0 },
        { paragraph: 1, fonts: [], colours: ['BED62F'], highlights: [], smallestHalfPoints: 16, pictures: 0 },
      ],
    }))
    expect(rules(f).sort()).toEqual(['brand/highlight', 'brand/lime-on-light', 'brand/off-palette', 'brand/small-text'])
    expect(f.find((x) => x.rule === 'brand/small-text')?.title).toBe('Text at 6 pt')
  })

  it('falls back to paragraph fonts when the run-level scan is off', () => {
    const f = brandFindings(snap([{ text: 'Red', font: { name: 'Aptos', size: 7, color: '#FF0000', highlight: 'Yellow', bold: false } }]))
    expect(rules(f).sort()).toEqual(['brand/highlight', 'brand/off-palette', 'brand/small-text'])
  })
})

describe('before you send', () => {
  it('lists tracked changes, comments, prompts, placeholders and a sticker', () => {
    const f = sendFindings(snap([], {
      trackedChanges: 2, comments: 1,
      controls: [
        { index: 0, tag: 'BwBGuidance', title: 'Writing prompt', text: 'Say what the report is for.' },
        { index: 1, tag: 'BwBPlaceholder', title: 'Client', text: '[Client name]' },
        { index: 2, tag: 'BwBPlaceholder', title: 'Year', text: '2026' },
        { index: 3, tag: 'BwBSticker', title: 'BwB sticker', text: 'DRAFT' },
      ],
    }))
    expect(f.map((x) => x.title)).toEqual([
      '2 tracked changes still in the document', '1 comment in the document', '1 writing prompt left in',
      'Not filled in: [Client name]', '“DRAFT” sticker in the header',
    ])
    expect(f.find((x) => x.rule === 'send/sticker')?.advisory).toBe(true)
  })

  it('checks the title, author and company properties, the footer and the file name', () => {
    const base = snap([])
    const f = sendFindings(snap([], {
      info: { ...base.info, title: '', author: 'jsmith', company: '' },
      headersFooters: [{ section: 0, kind: 'footer', text: 'Bankers without Boundaries', pictures: 0 }],
      fileName: 'Report - Copy (2).docx',
    }))
    expect(rules(f)).toEqual(expect.arrayContaining(['send/no-title', 'send/properties', 'send/page-numbers']))
    expect(f.find((x) => x.rule === 'send/properties')?.fix).toEqual({ kind: 'properties', values: { author: BWB, company: BWB } })
    expect(rules(f).some((r) => r.startsWith('send/file-name-'))).toBe(true)
    expect(sendFindings(base)).toEqual([])
  })
})

describe('consistency', () => {
  it('flags acronyms used before they are spelled out, or never spelled out', () => {
    const { findings } = consistencyFindings(snap([
      { text: 'The UNDP pilot worked.' },
      { text: 'The United Nations Development Programme (UNDP) funded it, with MDB support in the UK.' },
    ]))
    expect(findings.map((x) => x.title).sort()).toEqual(['MDB is never spelled out', 'UNDP is used before it is spelled out'])
  })

  it('asks for an abbreviations table in a long report with many acronyms', () => {
    const acronyms = ['GCF', 'MDB', 'NDC', 'SDG', 'IMF', 'OECD', 'UNEP', 'IFAD']
    const ps: P[] = acronyms.map((a) => ({ text: `${a} is used here.` }))
    ps.push({ text: Array.from({ length: 5000 }, () => 'word').join(' ') })
    expect(rules(consistencyFindings(snap(ps)).findings)).toContain('consistency/abbreviations-table')
    ps.unshift(h(1, 'Abbreviations and acronyms'))
    expect(rules(consistencyFindings(snap(ps)).findings)).not.toContain('consistency/abbreviations-table')
  })

  it('notes compounds written more than one way', () => {
    const { findings } = consistencyFindings(snap([
      { text: 'Decision-making is slow.' }, { text: 'Better decision-making helps.' }, { text: 'Their decision making improved.' },
    ]))
    const f = findings.find((x) => x.rule === 'consistency/hyphenation')
    expect(f?.advisory).toBe(true)
    expect(f?.detail).toContain('“decision-making”')
  })
})

describe('house-style wording', () => {
  it('turns wording issues into findings with fixes, and placeholders into major send items', () => {
    const f = writingFindings(snap([{ text: 'The organization will report by TBC.' }]))
    expect(f.find((x) => x.rule === 'writing/us-spelling')).toMatchObject({
      severity: 'copyedit', group: 'writing', fix: { kind: 'replace', replacement: 'organisation' }, fixLabel: 'Change to “organisation”',
    })
    expect(f.find((x) => x.rule === 'writing/placeholder')).toMatchObject({ severity: 'major', group: 'send' })
    expect(f.every((x) => !(x.detail ?? '').includes("'"))).toBe(true)
  })

  it('ignores the double space Word shows where an inline icon sits between two words', () => {
    const text = 'Water  security'
    expect(writingFindings(snap([{ text }])).map((x) => x.rule)).toEqual(['writing/double-space'])
    const withIcon = snap([{ text }], { formatting: [{ paragraph: 0, fonts: [], colours: [], highlights: [], smallestHalfPoints: null, pictures: 1 }] })
    expect(writingFindings(withIcon)).toEqual([])
  })

  it('leaves pull quotes and writing prompts word for word', () => {
    const f = writingFindings(snap([
      { text: 'We analyze everything \u2014 always.', style: 'BwB Pull Quote' },
      { text: 'Say what the program is for.' },
    ], { controls: [{ index: 0, tag: 'BwBGuidance', title: '', text: 'Say what the program is for.' }] }))
    expect(f).toEqual([])
  })
})
