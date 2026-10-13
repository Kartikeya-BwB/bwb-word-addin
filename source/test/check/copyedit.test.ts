import { describe, expect, it } from 'vitest'
import { copyeditFindings } from '../../src/check/rulesCopyedit.ts'
import type { DocSnapshot, ParagraphInfo } from '../../src/check/types.ts'

type P = Partial<ParagraphInfo> & { text: string }

function snap(paragraphs: P[]): DocSnapshot {
  return {
    paragraphs: paragraphs.map((p, index) => ({
      index, style: 'Normal', styleBuiltIn: 'Normal', isListItem: false, inTable: false,
      font: { name: 'Aptos', size: 10, color: '#000000', highlight: null, bold: false }, ...p,
    })),
    tables: [], pictures: [], controls: [], headersFooters: [], fieldCodes: [], comments: 0, trackedChanges: 0, formatting: null,
    info: { title: 'T', subject: '', author: '', company: '', docType: '', project: '', deliverable: '', version: '', funder: '', clientTemplate: false },
    fileName: null, bwbStyles: true,
  }
}

const rules = (text: string): string[] => copyeditFindings(snap([{ text }])).map((f) => f.rule)

describe('copyedit checks', () => {
  it('finds a word typed twice and removes the repeat', () => {
    const [f] = copyeditFindings(snap([{ text: 'The fund will will invest in the the region.' }]))
    expect(f?.title).toBe('“will” twice in a row')
    expect(f?.fix).toMatchObject({ kind: 'replace', match: ' will', replacement: '' })
    expect(rules('The the cat')).toEqual(['copyedit/repeated-word'])
    expect(rules('They had had enough.')).toEqual(['copyedit/repeated-word'])
    expect(copyeditFindings(snap([{ text: 'They had had enough.' }]))[0]?.advisory).toBe(true)
    expect(rules('Phase 2 2026')).toEqual([])
  })

  it('finds unmatched brackets and quotation marks, but not list markers', () => {
    expect(rules('The fund (launched in 2026 is open.')).toEqual(['copyedit/unmatched'])
    expect(rules('A closing bracket) without an opening one.')).toEqual(['copyedit/unmatched'])
    expect(rules('She said “yes.')).toEqual(['copyedit/unmatched'])
    expect(rules('a) the first point')).toEqual([])
    expect(rules('Balanced (as here) and “quoted”.')).toEqual([])
  })

  it('removes spaces before punctuation and doubled punctuation, but leaves ellipses', () => {
    const [space] = copyeditFindings(snap([{ text: 'Results are strong , as expected.' }]))
    expect(space?.rule).toBe('copyedit/space-before-punctuation')
    expect(space?.fix).toMatchObject({ match: ' ', replacement: '' })
    expect(rules('Done,, then more..')).toEqual(['copyedit/doubled-punctuation', 'copyedit/doubled-punctuation'])
    expect(rules('To be continued...')).toEqual([])
    expect(rules('The rate rose to .5 per cent')).toEqual([])
  })

  it('notes lists that end some items with a full stop and some without', () => {
    const items = (texts: string[]): P[] => texts.map((text) => ({ text, isListItem: true, style: 'BwB Bullet Green Square' }))
    const mixed = copyeditFindings(snap(items(['Raise capital.', 'Build the pipeline', 'Report on impact'])))
    expect(mixed.map((f) => f.rule)).toEqual(['copyedit/list-endings'])
    expect(mixed[0]?.excerpt).toBe('Raise capital.')
    expect(copyeditFindings(snap(items(['Raise capital', 'Build the pipeline', 'Report on impact'])))).toEqual([])
    expect(copyeditFindings(snap(items(['One.', 'Two'])))).toEqual([])
  })

  it('leaves pull quotes word for word', () => {
    expect(copyeditFindings(snap([{ text: 'We we said it twice (on purpose', style: 'BwB Pull Quote' }]))).toEqual([])
  })
})
