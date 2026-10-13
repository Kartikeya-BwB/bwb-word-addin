import { readFileSync } from 'node:fs'
import { describe, expect, it } from 'vitest'
import {
  curly, fix, fixWithRules, protectedText, quoteIssues, scan, scanWithRules, WritingFixError,
  type RuledIssue, type RuleId, type WritingContext, type WritingIssue,
} from './writing.ts'
import { cpLength } from './writingUnicode.ts'

// ------------------------------------------------------------------------------------------- fixtures
type FixKey = 'plain' | 'quotes' | 'fragment' | 'fragmentQuotes'
type Outcome = [string, number] | { error: string }
interface CtxResult {
  scan: Array<[string, string, string, string, boolean]>
  spans: Array<[string, number, number]>
  fix: Record<FixKey, Outcome>
  rules: Record<string, string | { error: string }> | null
}
interface Case { id: string; text: string; curly: [string, number]; protected?: string; results: CtxResult[] }
interface Expected { contexts: WritingContext[]; fixRules: RuleId[]; meta: Record<string, string>; cases: Case[] }
interface Corpus { generated: { texts: string[] }; handwritten: Array<{ group: string; texts: string[] }> }

const fixture = (name: string): string => readFileSync(new URL(`../../test/fixtures/${name}`, import.meta.url), 'utf8')
const expected = JSON.parse(fixture('writing-expected.json')) as Expected
const corpus = JSON.parse(fixture('writing-corpus.json')) as Corpus
const FIX_VARIANTS: Record<FixKey, { fragment: boolean; quotes: boolean }> = {
  plain: { fragment: false, quotes: false },
  quotes: { fragment: false, quotes: true },
  fragment: { fragment: true, quotes: false },
  fragmentQuotes: { fragment: true, quotes: true },
}
const ctxName = (ctx: WritingContext): string => JSON.stringify(ctx)

function runFix(text: string, ctx: WritingContext, opts: { fragment: boolean; quotes: boolean }): Outcome {
  try {
    const r = fix(text, ctx, opts)
    return [r.text, r.changes]
  } catch (e) {
    if (e instanceof WritingFixError) return { error: e.pythonError }
    throw e
  }
}

/** Applies replacements left to right, skipping any that overlaps one already applied. */
function applyReplacements(text: string, issues: readonly WritingIssue[]): string {
  return applyReplacementsCounting(text, issues).text
}

function applyReplacementsCounting(text: string, issues: readonly WritingIssue[]): { text: string; skipped: number } {
  const withFix = issues
    .map((issue, order) => ({ issue, order }))
    .filter((x) => x.issue.replacement !== undefined)
    .sort((a, b) => a.issue.start - b.issue.start || a.order - b.order)
  let out = ''
  let last = 0
  let skipped = 0
  for (const { issue } of withFix) {
    if (issue.start < last) {
      skipped++
      continue
    }
    out += text.slice(last, issue.start) + (issue.replacement ?? '')
    last = issue.end
  }
  return { text: out + text.slice(last), skipped }
}

/** Mismatches are collected and asserted at the end, so one failing case shows next to every other one. */
function check(label: string, each: (c: Case, ctx: WritingContext, r: CtxResult) => unknown[]): void {
  const failures: unknown[] = []
  let comparisons = 0
  for (const c of expected.cases) {
    expected.contexts.forEach((ctx, i) => {
      const r = c.results[i]
      if (!r) throw new Error(`missing result for ${c.id}`)
      comparisons++
      for (const f of each(c, ctx, r)) failures.push({ id: c.id, ctx: ctxName(ctx), text: c.text, ...Object(f) })
    })
  }
  expect(failures.slice(0, 20), `${failures.length} of ${comparisons} ${label} comparisons differ`).toEqual([])
}

// ------------------------------------------------------------------------------------------- differential tests
describe('differential test against bwb_lib/writing.py', () => {
  it('has the corpus the fixture was generated from', () => {
    const texts = [...corpus.handwritten.flatMap((g) => g.texts), ...corpus.generated.texts]
    expect(texts).toEqual(expected.cases.map((c) => c.text))
    expect(corpus.handwritten.flatMap((g) => g.texts).length).toBeGreaterThanOrEqual(400)
    expect(corpus.generated.texts.length).toBeGreaterThanOrEqual(300)
  })

  it('protects the same spans as _protect()', () => {
    const failures = expected.cases
      .filter((c) => protectedText(c.text) !== (c.protected ?? c.text))
      .map((c) => ({ id: c.id, ts: protectedText(c.text), py: c.protected ?? c.text }))
    expect(failures).toEqual([])
  })

  it('scan() returns the same issues in the same order (code, severity, message, excerpt, fixable)', () => {
    check('scan', (c, ctx, r) => {
      const ts = scan(c.text, ctx).map((i) => [i.code, i.severity, i.message, i.excerpt, i.fixable])
      return JSON.stringify(ts) === JSON.stringify(r.scan) ? [] : [{ ts, py: r.scan }]
    })
  })

  it('scan() finds each issue with the same rule at the same place as Python', () => {
    check('span', (c, ctx, r) => {
      const prot = protectedText(c.text)
      const cp = (offset: number): number => cpLength(prot.slice(0, offset))
      const ts = scanWithRules(c.text, ctx).map((x) => [x.rule, cp(x.protectedStart), cp(x.protectedEnd)])
      return JSON.stringify(ts) === JSON.stringify(r.spans) ? [] : [{ ts, py: r.spans }]
    })
  })

  for (const [key, opts] of Object.entries(FIX_VARIANTS)) {
    it(`fix() matches Python with fragment=${opts.fragment}, quotes=${opts.quotes}`, () => {
      check(`fix ${key}`, (c, ctx, r) => {
        const ts = runFix(c.text, ctx, opts)
        const py = r.fix[key as FixKey]
        return JSON.stringify(ts) === JSON.stringify(py) ? [] : [{ ts, py }]
      })
    })
  }

  it('curly() matches Python', () => {
    const failures = expected.cases
      .map((c) => ({ id: c.id, ts: curly(c.text), py: c.curly }))
      .filter((x) => x.ts.text !== x.py[0] || x.ts.changes !== x.py[1])
    expect(failures).toEqual([])
  })

  it('per-issue replacements of each rule reproduce Python fix() with only that rule switched on', () => {
    check('single-rule', (c, ctx, r) => {
      if (r.rules === null) return [] // the text holds placeholder characters; _restore() would interfere
      const issues = scanWithRules(c.text, ctx)
      const out: unknown[] = []
      for (const rule of expected.fixRules) {
        const own = issues.filter((x) => x.rule === rule).map((x) => x.issue)
        const failed = own.some((i) => i.fixable && i.replacement === undefined)
        const ts = failed ? { error: 'KeyError' } : applyReplacements(c.text, own)
        const py = r.rules[rule] ?? c.text
        if (JSON.stringify(ts) !== JSON.stringify(py)) out.push({ rule, ts, py })
      }
      return out
    })
  })
})

// ------------------------------------------------------------------------------------------- replacements vs fix()
/**
 * fix() runs the rules one after another, so a later rule sees the text an earlier rule produced, while each
 * replacement is computed on the original text. The two agree unless rules interact.
 */
describe('replacements compared with fix()', () => {
  it('applying all replacements of a scan equals fix() whenever only one rule is involved', () => {
    const fixRules = new Set<string>(expected.fixRules)
    let total = 0
    let equal = 0
    const interacting: string[] = []
    const failures: unknown[] = []
    for (const c of expected.cases) {
      if (c.text.includes('\uE000') || c.text.includes('\uE001')) continue
      for (const ctx of expected.contexts) {
        let traced: ReturnType<typeof fixWithRules>
        try {
          traced = fixWithRules(c.text, ctx)
        } catch {
          continue // Python raises for this text; nothing to compare
        }
        const issues = scanWithRules(c.text, ctx)
        const flagged = new Set(issues.filter((x) => fixRules.has(x.rule)).map((x) => x.rule))
        const changed = Object.keys(traced.changesByRule) as RuleId[]
        const involved = new Set<string>([...flagged, ...changed])
        const { text: replaced, skipped } = applyReplacementsCounting(c.text, issues.map((x) => x.issue))
        total++
        if (replaced === traced.text) {
          equal++
          continue
        }
        if (involved.size <= 1) {
          failures.push({ id: c.id, ctx: ctxName(ctx), fix: traced.text, replaced })
          continue
        }
        // Rules that changed more matches in fix() than they had replacements for in the original text.
        const own = (r: RuleId): RuledIssue[] => issues.filter((x) => x.rule === r)
        const extra = changed.filter((r) =>
          (traced.changesByRule[r] ?? 0) > own(r).filter((x) => x.issue.replacement !== undefined).length)
        const contextual = extra.filter((r) => own(r).some((x) => x.issue.replacement === undefined))
        const cascades = extra.filter((r) => !contextual.includes(r))
        const reason = cascades.length ? `new match for ${cascades.join(', ')} after an earlier fix`
          : contextual.length ? `${contextual.join(', ')} not fixable in the original (proper noun), fixable after an earlier fix`
            : skipped ? 'overlapping matches of two rules' : 'replacements interact'
        interacting.push(`${c.id} ${ctxName(ctx)}: ${reason}; fix() ${JSON.stringify(traced.text)}, ` +
          `replacements ${JSON.stringify(replaced)}`)
      }
    }
    console.info(`replacements equal fix() in ${equal} of ${total} case/context pairs ` +
      `(${((100 * equal) / total).toFixed(2)} per cent); ${interacting.length} differ, all with two or more rules involved`)
    for (const line of interacting) console.info(`  ${line}`)
    expect(failures).toEqual([])
  })

  it('with quoteIssues() added, the replacements equal fix() with quotes whenever only one rule is involved', () => {
    const fixRules = new Set<string>(expected.fixRules)
    let total = 0
    let equal = 0
    const differ: string[] = []
    const failures: unknown[] = []
    for (const c of expected.cases) {
      if (c.text.includes('\uE000') || c.text.includes('\uE001')) continue
      for (const ctx of expected.contexts) {
        let traced: ReturnType<typeof fixWithRules>
        try {
          traced = fixWithRules(c.text, ctx, { quotes: true })
        } catch {
          continue
        }
        const issues = scanWithRules(c.text, ctx)
        const quotes = quoteIssues(c.text)
        const involved = new Set<string>([
          ...issues.filter((x) => fixRules.has(x.rule)).map((x) => x.rule), ...Object.keys(traced.changesByRule),
          ...(quotes.length ? ['quotes'] : []),
        ])
        const replaced = applyReplacements(c.text, [...issues.map((x) => x.issue), ...quotes])
        total++
        if (replaced === traced.text) equal++
        else if (involved.size <= 1) failures.push({ id: c.id, ctx: ctxName(ctx), fix: traced.text, replaced })
        else differ.push(`${c.id} ${ctxName(ctx)} [${[...involved].join(', ')}]`)
      }
    }
    console.info(`with quotes: replacements equal fix() in ${equal} of ${total} case/context pairs ` +
      `(${((100 * equal) / total).toFixed(2)} per cent); ${differ.length} differ, all with two or more rules involved`)
    for (const line of differ) console.info(`  ${line}`)
    expect(failures).toEqual([])
  })
})

// ------------------------------------------------------------------------------------------- hand-checked cases
type Picked = Pick<WritingIssue, 'code' | 'start' | 'end' | 'match' | 'fixable'> & { replacement?: string }
const pick = (i: WritingIssue): Picked => {
  const out: Picked = { code: i.code, start: i.start, end: i.end, match: i.match, fixable: i.fixable }
  if (i.replacement !== undefined) out.replacement = i.replacement
  return out
}
const EM = '\u2014'
const SMILE = '\u{1F600}'

describe('issue offsets, matches and replacements', () => {
  const cases: Array<[string, WritingContext, Picked[]]> = [
    [`Prices rose ${EM} sharply.`, {}, [{ code: 'em-dash', start: 11, end: 14, match: ` ${EM} `, fixable: true, replacement: ', ' }]],
    [`${EM} leading`, {}, [{ code: 'em-dash', start: 0, end: 2, match: `${EM} `, fixable: true, replacement: '' }]],
    [`trailing ${EM}`, {}, [{ code: 'em-dash', start: 8, end: 10, match: ` ${EM}`, fixable: true, replacement: '' }]],
    [`2020${EM}2025`, {}, [{ code: 'em-dash', start: 4, end: 5, match: EM, fixable: true, replacement: ' to ' }]],
    [`\u00B2${EM}\u00B3`, {}, [{ code: 'em-dash', start: 1, end: 2, match: EM, fixable: true, replacement: ' to ' }]],
    [`The plan. ${EM} It works.`, {}, [{ code: 'em-dash', start: 9, end: 12, match: ` ${EM} `, fixable: true, replacement: ' ' }]],
    [`x.\uFEFF${EM} y`, {}, [{ code: 'em-dash', start: 3, end: 5, match: `${EM} `, fixable: true, replacement: ', ' }]],
    [`See https://x.com/a ${EM} now`, {}, [{ code: 'em-dash', start: 19, end: 22, match: ` ${EM} `, fixable: true, replacement: ', ' }]],
    [`${SMILE} color`, {}, [{ code: 'us-spelling', start: 3, end: 8, match: 'color', fixable: true, replacement: 'colour' }]],
    ['caf\u00E9 color', {}, [{ code: 'us-spelling', start: 5, end: 10, match: 'color', fixable: true, replacement: 'colour' }]],
    ['Z\u00FCrich center', {}, [{ code: 'us-spelling', start: 7, end: 13, match: 'center', fixable: true, replacement: 'centre' }]],
    ['Organizaci\u00F3n', {}, []],
    ['the World Health Organization', {}, [{ code: 'us-spelling', start: 17, end: 29, match: 'Organization', fixable: false }]],
    ['the World Health Organization', { title: true },
      [{ code: 'us-spelling', start: 17, end: 29, match: 'Organization', fixable: true, replacement: 'Organisation' }]],
    ['Organization matters.', {}, [{ code: 'us-spelling', start: 0, end: 12, match: 'Organization', fixable: true, replacement: 'Organisation' }]],
    ['Results from Analyze Inc.', {}, [{ code: 'us-spelling', start: 13, end: 20, match: 'Analyze', fixable: false }]],
    ['COLOR', {}, [{ code: 'us-spelling', start: 0, end: 5, match: 'COLOR', fixable: true, replacement: 'COLOUR' }]],
    ['Behaviors change.', {}, [{ code: 'us-spelling', start: 0, end: 9, match: 'Behaviors', fixable: true, replacement: 'Behaviours' }]],
    ['ANALYZING', {}, [{ code: 'us-spelling', start: 0, end: 9, match: 'ANALYZING', fixable: true, replacement: 'ANALYSING' }]],
    ['USD5m', {}, [{ code: 'currency-space', start: 0, end: 3, match: 'USD', fixable: true, replacement: 'USD ' }]],
    ['\u00A3 5m', {}, [{ code: 'currency', start: 0, end: 2, match: '\u00A3 ', fixable: true, replacement: 'GBP ' }]],
    ['Costs of \u20AC3m', {}, [{ code: 'currency', start: 9, end: 10, match: '\u20AC', fixable: true, replacement: 'EUR ' }]],
    ['US$5', {}, [{ code: 'currency-symbol', start: 0, end: 3, match: 'US$', fixable: false }]],
    ['January 5, 2024', {}, [{ code: 'date-format', start: 0, end: 15, match: 'January 5, 2024', fixable: true, replacement: '5 January 2024' }]],
    ['the 2nd of June 2026', {}, [{ code: 'date-format', start: 4, end: 20, match: '2nd of June 2026', fixable: true, replacement: '2 June 2026' }]],
    ['2024-01-05', {}, [{ code: 'date-format', start: 0, end: 10, match: '2024-01-05', fixable: true, replacement: '5 January 2024' }]],
    ['March \u0663, 2024', {}, [{ code: 'date-format', start: 0, end: 13, match: 'March \u0663, 2024', fixable: true, replacement: '3 March 2024' }]],
    ['5 Jan 2024', {}, [
      { code: 'date-format', start: 0, end: 10, match: '5 Jan 2024', fixable: true, replacement: '5 January 2024' },
      { code: 'date-format', start: 2, end: 10, match: 'Jan 2024', fixable: true, replacement: 'January 2024' },
    ]],
    ['5 Jan 2024', { compact: true }, [
      { code: 'date-format', start: 0, end: 10, match: '5 Jan 2024', fixable: false },
      { code: 'date-format', start: 2, end: 10, match: 'Jan 2024', fixable: false },
    ]],
    ['Percent', {}, [{ code: 'per-cent', start: 0, end: 7, match: 'Percent', fixable: true, replacement: 'Per cent' }]],
    ['PERCENT', {}, [{ code: 'per-cent', start: 0, end: 7, match: 'PERCENT', fixable: true, replacement: 'PER CENT' }]],
    ['2022 - 2025', {}, [{ code: 'range-hyphen', start: 0, end: 11, match: '2022 - 2025', fixable: true, replacement: '2022 to 2025' }]],
    ['5\u201310 per cent', {}, [
      { code: 'en-dash-range', start: 1, end: 2, match: '\u2013', fixable: true, replacement: ' to ' },
      { code: 'en-dash-compound', start: 1, end: 2, match: '\u2013', fixable: false },
    ]],
    ['Paris \u2013 London', {}, [{ code: 'en-dash', start: 5, end: 8, match: ' \u2013 ', fixable: true, replacement: ', ' }]],
    ['North\u2013South', {}, [{ code: 'en-dash-compound', start: 5, end: 6, match: '\u2013', fixable: false }]],
    ['the bankers without boundaries team', {},
      [{ code: 'brand-name', start: 4, end: 30, match: 'bankers without boundaries', fixable: true, replacement: 'Bankers without Boundaries' }]],
    ['BWB and info@bwb.earth', {}, [{ code: 'brand-name', start: 0, end: 3, match: 'BWB', fixable: true, replacement: 'BwB' }]],
    ['a  b', {}, [{ code: 'double-space', start: 1, end: 3, match: '  ', fixable: true, replacement: ' ' }]],
    ['{icon:leaf}  Text', {}, [{ code: 'double-space', start: 11, end: 13, match: '  ', fixable: true, replacement: ' ' }]],
    ['a \u00B7  b', {}, []],
    ['F\u0130BER optics', {}, [{ code: 'us-spelling', start: 0, end: 5, match: 'F\u0130BER', fixable: true }]],
    ['<<see `code` here>>', {}, [{ code: 'placeholder', start: 0, end: 19, match: '<<see `code` here>>', fixable: false }]],
    ['Let us touch base.', {}, [{ code: 'filler', start: 7, end: 17, match: 'touch base', fixable: false }]],
    ['https://x.com/color-center', {}, []],
  ]
  it.each(cases)('scan(%j, %j)', (text, ctx, want) => {
    const got = scan(text, ctx)
    expect(got.map(pick)).toEqual(want)
    for (const i of got) expect(text.slice(i.start, i.end)).toBe(i.match)
  })

  it('cuts excerpts by code points and leaves placeholders in them, as Python does', () => {
    const text = `${SMILE.repeat(30)} color ${SMILE.repeat(30)}`
    const [issue] = scan(text)
    expect(issue?.excerpt).toBe(`${SMILE.repeat(24)} color ${SMILE.repeat(24)}`)
    expect(scan('https://x.com color')[0]?.excerpt).toBe('\uE0000\uE001 color')
    expect(scan('<<see `code` here>>')[0]?.message).toBe("Placeholder text left in: '<<see \uE0000\uE001 here>>'")
  })

  it('lists straight quotes outside protected spans with their typographic replacement', () => {
    const text = '"Hello," she said. It\'s https://x.com/"q"'
    expect(quoteIssues(text).map(pick)).toEqual([
      { code: 'quotes', start: 0, end: 1, match: '"', fixable: true, replacement: '\u201C' },
      { code: 'quotes', start: 7, end: 8, match: '"', fixable: true, replacement: '\u201D' },
      { code: 'quotes', start: 21, end: 22, match: "'", fixable: true, replacement: '\u2019' },
    ])
    expect(quoteIssues("'90s")[0]?.replacement).toBe('\u2018')
    expect(quoteIssues(`${EM}"x"`)[0]?.replacement).toBe('\u201C')
    expect(quoteIssues('`it\'s`')).toEqual([])
    const [issue] = quoteIssues('He said "no".')
    expect(issue).toMatchObject({ severity: 'info', message: 'Straight quotation mark; use a typographic one', excerpt: 'He said "no".' })
  })
})

describe('fix() options and failures', () => {
  it('leaves boundary em dashes alone in fragments', () => {
    expect(fix(`${EM} a`)).toEqual({ text: 'a', changes: 1 })
    expect(fix(`${EM} a`, {}, { fragment: true })).toEqual({ text: `${EM} a`, changes: 0 })
    expect(fix(`a ${EM} b`, {}, { fragment: true })).toEqual({ text: 'a, b', changes: 1 })
  })

  it('sets typographic quotes only when asked, outside protected spans', () => {
    expect(fix('"a" it\'s')).toEqual({ text: '"a" it\'s', changes: 0 })
    expect(fix('"a" it\'s `b\'s`', {}, { quotes: true })).toEqual({ text: '\u201Ca\u201D it\u2019s `b\'s`', changes: 3 })
  })

  it('throws where Python raises', () => {
    expect(() => fix('F\u0130BER')).toThrow(WritingFixError)
    expect(() => fix('\uE0000\uE001 color')).toThrow(expect.objectContaining({ pythonError: 'IndexError' }))
    expect(fix('https://a.b \uE0000\uE001').text).toBe('https://a.b https://a.b')
  })

  it('counts the changes of each rule', () => {
    expect(fixWithRules(`Hi. ${EM} Color`)).toEqual({ text: 'Hi. Colour', changes: 2, changesByRule: { EM_DASH: 1, OUR_RE: 1 } })
  })
})
