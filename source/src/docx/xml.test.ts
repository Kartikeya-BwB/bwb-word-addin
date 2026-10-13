import { describe, expect, it } from 'vitest'
import {
  cpHead,
  cpLen,
  esc,
  escapeInline,
  field,
  para,
  ppr,
  pyMax,
  pyReSplitWs,
  pyRound,
  pyRoundTo,
  pySplit,
  pyStrip,
  pyStripChars,
  pySum,
  replaceFirst,
  rpr,
  run,
  spacer,
  splitInline,
} from './xml.ts'

describe('Python number semantics', () => {
  it('rounds halves to even like round(x)', () => {
    const cases: Array<[number, number]> = [
      [0.5, 0], [1.5, 2], [2.5, 2], [-0.5, 0], [-1.5, -2], [2.4999999999999996, 2],
      [1000000000000000.5, 1000000000000000], [8.746600566572239, 9], [16.997167138810198, 17],
    ]
    for (const [x, want] of cases) expect(pyRound(x) + 0, String(x)).toBe(want)
  })

  it('rounds to decimals on the exact binary value like round(x, 2)', () => {
    const cases: Array<[number, number]> = [
      [0.353, 0.35], [0.125, 0.12], [0.375, 0.38], [0.015, 0.01], [0.005, 0.01], [9 * 0.0353, 0.32],
      [10 * 0.0353, 0.35], [11 * 0.0353, 0.39], [12 * 0.0353, 0.42], [18 * 0.0353, 0.64], [40 * 0.0353, 1.41],
      [2.675, 2.67], [1.005, 1.0], [0.0353 * 14, 0.49],
    ]
    for (const [x, want] of cases) expect(pyRoundTo(x, 2), String(x)).toBe(want)
  })

  it('sums like Python 3.12+ sum(): compensated floats, plain ints after the first float', () => {
    const f = (v: number) => ({ v, isInt: false })
    const i = (v: number) => ({ v, isInt: true })
    expect(pySum(Array(10).fill(f(0.1))).v).toBe(1.0)
    expect(pySum([f(1e16), f(1.0), f(-1e16)]).v).toBe(1.0)
    expect(pySum([f(1e16), i(1), f(-1e16)]).v).toBe(0.0)
    expect(pySum([f(0.1), f(0.2), f(0.3)]).v).toBe(0.6)
    expect(pySum([i(4), f(9.2), i(13), f(0.7)]).v).toBe(26.9)
    expect(pySum([i(1), i(2), i(3)])).toEqual({ v: 6, isInt: true })
    expect(pySum([f(7.237624155400388), i(4), f(12.172184414459759), i(5)]).v).toBe(28.409808569860147)
    expect(pySum([])).toEqual({ v: 0, isInt: true })
  })

  it('keeps the first of equal values in max()', () => {
    expect(pyMax({ v: 4, isInt: true }, { v: 4, isInt: false })).toEqual({ v: 4, isInt: true })
    expect(pyMax({ v: 4, isInt: true }, { v: 4.5, isInt: false }, { v: 4.5, isInt: true })).toEqual({ v: 4.5, isInt: false })
  })
})

describe('Python string semantics', () => {
  it('counts and slices by code point', () => {
    expect(cpLen('a\u{1F30D}b')).toBe(3)
    expect(cpHead('\u{1F30D}\u{1F30D}abc', 3)).toBe('\u{1F30D}\u{1F30D}a')
    expect(cpHead('abc', 10)).toBe('abc')
  })

  it("uses Python's whitespace set for strip and split", () => {
    expect(pyStrip('\u{A0}\u{3000} x \u{2028}')).toBe('x')
    expect(pyStrip('\u{FEFF}x')).toBe('\u{FEFF}x')
    expect(pySplit(' a\u{A0}b  c\x1fd ')).toEqual(['a', 'b', 'c', 'd'])
    expect(pyReSplitWs(' a b ')).toEqual(['', 'a', 'b', ''])
    expect(pyStripChars('"quoted"', '"')).toBe('quoted')
  })

  it('replaces the first occurrence without $ patterns', () => {
    expect(replaceFirst('a-b-b', 'b', '$&$1')).toBe('a-$&$1-b')
  })
})

describe('XML helpers', () => {
  it('escapes like xml.sax.saxutils.escape with quotes, and drops characters XML cannot hold', () => {
    expect(esc('a & b < c > "d" \'e\'')).toBe('a &amp; b &lt; c &gt; &quot;d&quot; \'e\'')
    expect(esc('bell\u0007 tab\t vt\u000b')).toBe('bell tab\t vt')
    expect(esc(12)).toBe('12')
  })

  it('writes run properties in schema order and sizes in half points', () => {
    expect(rpr({ vert: 'superscript', u: 'single', sz: 9, color: 'FFFFFF', noproof: true, caps: true, i: true, b: false, font: 'Aptos', rstyle: 'X' })).toBe(
      '<w:rPr><w:rStyle w:val="X"/><w:rFonts w:ascii="Aptos" w:eastAsia="Aptos" w:hAnsi="Aptos" w:cs="Aptos"/>' +
        '<w:b w:val="0"/><w:bCs w:val="0"/><w:i/><w:iCs/><w:caps/><w:noProof/><w:color w:val="FFFFFF"/>' +
        '<w:sz w:val="18"/><w:szCs w:val="18"/><w:u w:val="single"/><w:vertAlign w:val="superscript"/></w:rPr>',
    )
    expect(rpr({ sz: 10.25 })).toContain('w:val="20"')
    expect(rpr({ sz: 10.75 })).toContain('w:val="22"')
    expect(rpr({})).toBe('')
  })

  it('splits tabs and line breaks into their own runs', () => {
    expect(run('a\tb\nc', { b: true })).toBe(
      '<w:r><w:rPr><w:b/><w:bCs/></w:rPr><w:t xml:space="preserve">a</w:t></w:r><w:r><w:rPr><w:b/><w:bCs/></w:rPr><w:tab/></w:r>' +
        '<w:r><w:rPr><w:b/><w:bCs/></w:rPr><w:t xml:space="preserve">b</w:t></w:r><w:r><w:rPr><w:b/><w:bCs/></w:rPr><w:br/></w:r>' +
        '<w:r><w:rPr><w:b/><w:bCs/></w:rPr><w:t xml:space="preserve">c</w:t></w:r>',
    )
    expect(run('')).toBe('')
  })

  it('writes a complex field with its cached result', () => {
    expect(field('PAGE', '1')).toBe(
      '<w:r><w:fldChar w:fldCharType="begin"/></w:r><w:r><w:instrText xml:space="preserve"> PAGE </w:instrText></w:r>' +
        '<w:r><w:fldChar w:fldCharType="separate"/></w:r><w:r><w:t>1</w:t></w:r><w:r><w:fldChar w:fldCharType="end"/></w:r>',
    )
  })

  it('orders paragraph properties by the schema, whatever the argument order', () => {
    expect(
      ppr({
        sect: '<w:sectPr/>',
        jc: 'left',
        ind: { left: 216, right: 216 },
        spacing: { before: 120, after: 120 },
        tabs: ['<w:tab w:val="right" w:pos="9458"/>'],
        extra: '<w:pBdr><w:left w:val="single"/></w:pBdr>',
        num: [2, 1],
        page_break_before: true,
        keep_lines: true,
        keep_next: true,
        style: 'S',
      }),
    ).toBe(
      '<w:pPr><w:pStyle w:val="S"/><w:keepNext/><w:keepLines/><w:pageBreakBefore/><w:numPr><w:ilvl w:val="1"/><w:numId w:val="2"/></w:numPr>' +
        '<w:pBdr><w:left w:val="single"/></w:pBdr><w:tabs><w:tab w:val="right" w:pos="9458"/></w:tabs>' +
        '<w:spacing w:before="120" w:after="120"/><w:ind w:left="216" w:right="216"/><w:jc w:val="left"/><w:sectPr/></w:pPr>',
    )
    expect(ppr({ extra: '<w:outlineLvl w:val="0"/>', jc: 'both' })).toBe('<w:pPr><w:jc w:val="both"/><w:outlineLvl w:val="0"/></w:pPr>')
    expect(para()).toBe('<w:p></w:p>')
    expect(spacer(120)).toBe(
      '<w:p><w:pPr><w:spacing w:before="0" w:after="0" w:line="120" w:lineRule="exact"/><w:rPr><w:sz w:val="2"/><w:szCs w:val="2"/></w:rPr></w:pPr></w:p>',
    )
  })
})

describe('inline mark-up', () => {
  it('splits bold, italic, links and footnotes', () => {
    expect(splitInline('a **b** *c* [d](https://e.org) ^[f *g*] h')).toEqual([
      { kind: 'text', value: 'a ', extra: null },
      { kind: 'bold', value: 'b', extra: null },
      { kind: 'text', value: ' ', extra: null },
      { kind: 'italic', value: 'c', extra: null },
      { kind: 'text', value: ' ', extra: null },
      { kind: 'link', value: 'd', extra: 'https://e.org' },
      { kind: 'text', value: ' ', extra: null },
      { kind: 'footnote', value: 'f *g*', extra: null },
      { kind: 'text', value: ' h', extra: null },
    ])
  })

  it('honours backslash escapes', () => {
    expect(splitInline('\\*not\\* \\[x\\](y) \\^[z] \\\\')).toEqual([{ kind: 'text', value: '*not* [x](y) ^[z] \\', extra: null }])
  })

  it('escapes text so that it reads back unchanged', () => {
    const s = 'a*b [c](d) ^[e] \\ f'
    expect(splitInline(escapeInline(s))).toEqual([{ kind: 'text', value: s, extra: null }])
  })

  it('treats a lone asterisk and an unclosed bracket as text', () => {
    expect(splitInline('5 * 3 and [open')).toEqual([{ kind: 'text', value: '5 * 3 and [open', extra: null }])
  })
})
