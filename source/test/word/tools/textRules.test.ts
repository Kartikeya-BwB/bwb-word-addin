import { describe, expect, it } from 'vitest'
import { captionKind, captionLabel, captionTitle } from '../../../src/word/tools/common.ts'
import { fragmentRisk, isRemovableEmpty, stripRunFormatting } from '../../../src/word/tools/ooxml.ts'
import { SYMBOLS, applyJoins, joinAction, nonBreakingSpans, planJoins } from '../../../src/word/tools/textRules.ts'

const LONG = 'Climate finance flows to sub-Saharan Africa grew in 2025 as the'

describe('joining lines pasted from a PDF', () => {
  it('joins a line that ends mid-sentence with a space', () => {
    expect(applyJoins([LONG, 'GCF scaled up its pipeline.'])).toEqual([`${LONG} GCF scaled up its pipeline.`])
  })

  it('keeps breaks after a full stop, colon or semicolon, before a list item, and around blank lines', () => {
    const text = ['The fund has three windows:', '\u2022 grants for early-stage projects and the', 'technical assistance they need;', '', 'Next paragraph.']
    expect(applyJoins(text)).toEqual([
      'The fund has three windows:',
      '\u2022 grants for early-stage projects and the technical assistance they need;',
      '',
      'Next paragraph.',
    ])
  })

  it('rejoins a word hyphenated at the end of a line, but keeps the hyphen before a capital or a digit', () => {
    expect(applyJoins([`${LONG} finan-`, 'cing gap widened.'])).toEqual([`${LONG} financing gap widened.`])
    expect(applyJoins([`${LONG} COVID-`, '19 recovery.'])).toEqual([`${LONG} COVID-19 recovery.`])
  })

  it('treats a short line as a heading or the end of a paragraph', () => {
    expect(applyJoins(['Introduction', `${LONG} market.`])).toEqual(['Introduction', `${LONG} market.`])
  })

  it('joins manual line breaks inside a paragraph too, and collapses double spaces', () => {
    expect(applyJoins([`${LONG} \u000bfund  grew.`])).toEqual([`${LONG} fund grew.`])
  })

  it('numbers manual line breaks and those after a hyphen, for matching Word\u2019s search hits', () => {
    const plan = planJoins([`${LONG} finan-\u000bcing ${LONG}\u000bfund grew`, 'next line'])
    expect(plan.map((b) => [b.kind, b.index, b.hyphenIndex, b.action])).toEqual([
      ['line', 0, 0, 'dehyphenate'],
      ['line', 1, -1, 'space'],
      ['paragraph', -1, -1, 'keep'],
    ])
  })

  it('decides each break from the two lines either side', () => {
    expect(joinAction('the end.', 'Next', 10)).toBe('keep')
    expect(joinAction('as the ', 'fund', 10)).toBe('join')
    expect(joinAction('as the', '1. First', 10)).toBe('keep')
  })
})

describe('nonBreakingSpans', () => {
  const apply = (text: string): string => {
    let out = text
    for (const s of [...nonBreakingSpans(text)].reverse()) out = out.slice(0, s.start) + s.replacement + out.slice(s.end)
    return out
  }

  it('keeps a currency code with its amount', () => {
    expect(apply('It raised USD 5m and EUR (2.1bn) in 2025.')).toBe('It raised USD\u00a05m and EUR\u00a0(2.1bn) in 2025.')
  })

  it('keeps a number with per cent', () => {
    expect(apply('Up 12 per cent on 2024.')).toBe('Up 12\u00a0per\u00a0cent on 2024.')
  })

  it('keeps a date in D Month YYYY form together', () => {
    expect(apply('Due on 5 October 2026 or 15 Sept. 2027.')).toBe('Due on 5\u00a0October\u00a02026 or 15\u00a0Sept.\u00a02027.')
  })

  it('keeps a number with its unit, but not with an ordinary word', () => {
    expect(apply('A 20 MW plant on 5 km of land, 3 men and 4 tonnes.')).toBe('A 20\u00a0MW plant on 5\u00a0km of land, 3 men and 4 tonnes.')
    expect(apply('USD 2.5 million')).toBe('USD\u00a02.5\u00a0million')
  })

  it('gives each span the words around its spaces, so a search finds that place', () => {
    expect(nonBreakingSpans('Costs of USD 5m rose 12 per cent.')).toEqual([
      { start: 9, end: 15, replacement: 'USD\u00a05m' },
      { start: 21, end: 33, replacement: '12\u00a0per\u00a0cent.' },
    ])
  })

  it('leaves text without figures alone', () => {
    expect(nonBreakingSpans('Nothing to keep together here.')).toEqual([])
  })
})

describe('the symbols palette', () => {
  it('has the BwB symbols, written as escapes in the source', () => {
    expect(SYMBOLS.map((s) => s.char).join('')).toBe('\u2713\u2717\u2192\u2190\u2191\u2193\u2022\u00b7\u00b0\u20ac\u00a3\u00b1\u00d7\u2264\u2265\u2248\u2026\u2018\u2019\u201c\u201d')
  })
})

const pkg = (body: string): string => `<pkg:package><pkg:part pkg:name="/word/document.xml"><pkg:xmlData><w:document><w:body>${body}<w:sectPr/></w:body></w:document></pkg:xmlData></pkg:part></pkg:package>`

describe('stripRunFormatting', () => {
  it('removes fonts, sizes, colours, highlighting and text effects, and keeps emphasis, styles and superscript', () => {
    const run = '<w:r w:rsidR="00A1"><w:rPr><w:rStyle w:val="Hyperlink"/><w:rFonts w:ascii="Arial"/><w:b/><w:i/><w:u w:val="single"/>' +
      '<w:color w:val="FF0000"/><w:sz w:val="28"/><w:highlight w:val="yellow"/><w:vertAlign w:val="superscript"/>' +
      '<w14:textFill><w14:solidFill><w14:srgbClr w14:val="FF0000"/></w14:solidFill></w14:textFill></w:rPr><w:t>x</w:t></w:r>'
    const { xml, removed } = stripRunFormatting(pkg(`<w:p><w:pPr><w:rPr><w:sz w:val="40"/></w:rPr></w:pPr>${run}</w:p>`))
    expect(removed).toBe(5)
    expect(xml).toContain('<w:rPr><w:rStyle w:val="Hyperlink"/><w:b/><w:i/><w:u w:val="single"/><w:vertAlign w:val="superscript"/></w:rPr><w:t>x</w:t>')
    expect(xml).toContain('<w:pPr><w:rPr><w:sz w:val="40"/></w:rPr></w:pPr>')
  })

  it('drops run properties that hold nothing to keep, and leaves plain runs alone', () => {
    const { xml, removed } = stripRunFormatting('<w:p><w:r><w:rPr><w:rFonts w:ascii="Calibri"/></w:rPr><w:t>a</w:t></w:r><w:r><w:t>b</w:t></w:r></w:p>')
    expect(removed).toBe(1)
    expect(xml).toBe('<w:p><w:r><w:t>a</w:t></w:r><w:r><w:t>b</w:t></w:r></w:p>')
  })
})

describe('fragmentRisk and isRemovableEmpty', () => {
  it('flags fragments that rewriting could damage', () => {
    expect(fragmentRisk(pkg('<w:p><w:ins w:id="1"><w:r><w:t>new</w:t></w:r></w:ins></w:p>'))).toBe('tracked changes')
    expect(fragmentRisk(pkg('<w:p><w:commentRangeStart w:id="0"/><w:r><w:t>x</w:t></w:r></w:p>'))).toBe('comments')
    expect(fragmentRisk(pkg('<w:p><w:bookmarkStart w:id="0" w:name="_Ref123"/></w:p>'))).toBe('bookmarks')
    expect(fragmentRisk(pkg('<w:p><w:bookmarkStart w:id="0" w:name="_GoBack"/><w:r><w:t>x</w:t></w:r></w:p>'))).toBeNull()
    expect(fragmentRisk(pkg('<w:tbl><w:tr><w:tc><w:p/></w:tc></w:tr></w:tbl>'))).toBe('table layout')
  })

  it('tells a truly empty paragraph from one holding a break, a picture or a section break', () => {
    expect(isRemovableEmpty(pkg('<w:p><w:pPr><w:pStyle w:val="Normal"/></w:pPr></w:p>'))).toBe(true)
    expect(isRemovableEmpty(pkg('<w:p><w:r><w:t xml:space="preserve">  </w:t></w:r></w:p>'))).toBe(true)
    expect(isRemovableEmpty(pkg('<w:p><w:r><w:br w:type="page"/></w:r></w:p>'))).toBe(false)
    expect(isRemovableEmpty(pkg('<w:p><w:r><w:drawing/></w:r></w:p>'))).toBe(false)
    expect(isRemovableEmpty(pkg('<w:p><w:pPr><w:sectPr/></w:pPr></w:p>'))).toBe(false)
    expect(isRemovableEmpty(pkg('<w:p><w:bookmarkStart w:id="1" w:name="BwB_Annex"/></w:p>'))).toBe(false)
  })
})

describe('captions', () => {
  it('knows a caption by its BwB style or its label, but not a sentence that starts with one', () => {
    expect(captionKind({ text: 'Anything', style: 'BwB Image Caption' })).toBe('Figure')
    expect(captionKind({ text: 'Anything', style: 'BwB Table Caption' })).toBe('Table')
    expect(captionKind({ text: 'Figure 3. Flows', style: 'Normal' })).toBe('Figure')
    expect(captionKind({ text: 'Table 2: Sizes', style: 'Normal' })).toBe('Table')
    expect(captionKind({ text: 'Figure 3 shows the flows.', style: 'Normal' })).toBeNull()
    expect(captionKind({ text: 'Flows', style: 'Caption', styleBuiltIn: 'Caption' })).toBe('Figure')
    expect(captionLabel('Table 3. Green bonds')).toBe('Table 3')
    expect(captionTitle('Figure 2. Climate finance flows, 2025. Source: OECD (2026)')).toBe('Climate finance flows, 2025.')
    expect(captionTitle('Figure 3. [Title] Source: [Source]')).toBe('')
  })
})
