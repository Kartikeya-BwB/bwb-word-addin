/**
 * BwB house writing rules: detection and safe automatic fixes. A line-by-line port of bwb_lib/writing.py from the
 * bwb-document-format skill. For the same input, scan(), fix() and curly() return what the Python returns; the
 * differential test (writing.test.ts) checks this against outputs recorded from Python by
 * scripts/gen-writing-fixtures.py.
 *
 * Additions for Word: every issue carries its offsets in the original text, the matched text and, when fixable, the
 * text fix() would write for that match if no other rule had run first (`replacement`). quoteIssues() lists straight
 * quotation marks as issues, which the Python only fixes.
 *
 * Known divergences from Python, kept deliberately small:
 * - Unicode version. The character classes use JavaScript's \p{...} properties, which follow the engine's Unicode
 *   version (16.0 in Node 22), while Python 3.13 follows Unicode 15.1. Results can differ only for characters that are
 *   unassigned in the Python's Unicode database, such as the scripts added in Unicode 16; writingUnicode.test.ts checks
 *   every code point against tables recorded from Python.
 * - Exceptions. Where Python's fix() raises, fix() throws a WritingFixError naming the Python exception: KeyError for a
 *   US spelling written with U+0130, U+0131 or U+017F (FIBER with a dotted capital I), IndexError or ValueError when
 *   the text itself holds the private-use placeholder characters U+E000 and U+E001. scan() never throws; such issues
 *   have no `replacement`.
 * - Lone surrogates, which Word does not produce, are treated as single characters as Python would, but are untested.
 */
import * as P from './writingPatterns.ts'
import { RULES, WritingFixError, type RuleId, type WritingCode, type WritingContext, type WritingSeverity } from './writingRules.ts'
import { codePointAt, codePointBefore, cpLength, moveByCodePoints, pyInt, pyIsAlnum } from './writingUnicode.ts'

export { WritingFixError } from './writingRules.ts'
export type { PythonError, RuleId, WritingCode, WritingContext, WritingSeverity } from './writingRules.ts'

export interface WritingIssue {
  code: WritingCode
  severity: WritingSeverity
  /** Identical to Python. */
  message: string
  /** Identical to Python: 25 characters either side of the match in the protected text, newlines as spaces. */
  excerpt: string
  /** Identical to Python. */
  fixable: boolean
  /** UTF-16 offsets of the match in the original text. */
  start: number
  end: number
  /** text.slice(start, end) */
  match: string
  /** Fixable issues only: what fix() writes in place of this match when this rule runs on the original text. */
  replacement?: string
}
export interface FixOptions { fragment?: boolean; quotes?: boolean }

// ------------------------------------------------------------------------------------------- protect and restore
interface Hole { pStart: number; pEnd: number; oStart: number; oEnd: number }
interface Protected { s: string; holes: string[]; spans: Hole[] }

/** _protect(): hides URLs, e-mail addresses, code spans, link targets and icon tokens behind U+E000 n U+E001. */
function protect(text: string): Protected {
  const holes: string[] = []
  const spans: Hole[] = []
  let s = ''
  let last = 0
  for (const m of text.matchAll(P.protectPattern(text))) {
    const marker = `${P.HOLE_OPEN}${holes.length}${P.HOLE_CLOSE}`
    s += text.slice(last, m.index)
    spans.push({ pStart: s.length, pEnd: s.length + marker.length, oStart: m.index, oEnd: m.index + m[0].length })
    s += marker
    holes.push(m[0])
    last = m.index + m[0].length
  }
  return { s: s + text.slice(last), holes, spans }
}

/** The text the rules run on: protected spans replaced by U+E000 n U+E001 placeholders. Excerpts are cut from it. */
export function protectedText(text: string): string {
  return protect(text).s
}

/** _restore(), including its failures when the text already contained placeholder characters. */
function restore(s: string, holes: readonly string[]): string {
  return s.replace(P.RESTORE, (_whole: string, digits: string) => {
    // int() refuses strings of more than 4300 digits (sys.get_int_max_str_digits()).
    if (cpLength(digits) > 4300) throw new WritingFixError('ValueError', 'Exceeds the limit for integer string conversion')
    let index = 0
    for (const ch of digits) {
      index = index * 10 + pyInt(ch)
      if (index >= holes.length) break
    }
    const hole = index < holes.length ? holes[index] : undefined
    if (hole === undefined) throw new WritingFixError('IndexError', 'list index out of range')
    return hole
  })
}

/** Maps an offset in the protected text to the original text; offsets inside a placeholder go to its edges. */
function toOriginal(p: Protected, pos: number, edge: 'start' | 'end'): number {
  let delta = 0
  for (const h of p.spans) {
    if (pos <= h.pStart) break
    if (pos < h.pEnd) return edge === 'start' ? h.oStart : h.oEnd
    delta += h.oEnd - h.oStart - (h.pEnd - h.pStart)
  }
  return pos + delta
}

function excerptOf(s: string, start: number, end: number): string {
  return s.slice(moveByCodePoints(s, start, -25), moveByCodePoints(s, end, 25)).replaceAll('\n', ' ')
}

// ------------------------------------------------------------------------------------------- scan and fix
/** An issue with the rule that found it and its UTF-16 span in the protected text (for tests and tooling). */
export interface RuledIssue { rule: RuleId; issue: WritingIssue; protectedStart: number; protectedEnd: number }

/** scan(), keeping for each issue the rule that produced it. */
export function scanWithRules(text: string, ctx: WritingContext = {}): RuledIssue[] {
  const p = protect(text)
  const s = p.s
  const out: RuledIssue[] = []
  for (const rule of RULES) {
    for (const m of s.matchAll(rule.re)) {
      if (rule.skip?.(m[0])) continue
      const pStart = m.index
      const pEnd = m.index + m[0].length
      const start = toOriginal(p, pStart, 'start')
      const end = toOriginal(p, pEnd, 'end')
      const fixable = rule.fixable(s, m, ctx)
      const issue: WritingIssue = {
        code: rule.code, severity: rule.severity(ctx), message: rule.message(m[0]), excerpt: excerptOf(s, pStart, pEnd),
        fixable, start, end, match: text.slice(start, end),
      }
      if (fixable && rule.substitute) {
        try {
          issue.replacement = rule.substitute(m, s, ctx, false)
        } catch (e) {
          if (!(e instanceof WritingFixError)) throw e
        }
      }
      out.push({ rule: rule.id, issue, protectedStart: pStart, protectedEnd: pEnd })
    }
  }
  return out
}

/** scan(): every issue in one text, in the order of the Python. */
export function scan(text: string, ctx: WritingContext = {}): WritingIssue[] {
  return scanWithRules(text, ctx).map((r) => r.issue)
}

/** re.sub with a function, counting the matches whose replacement differs from the match. */
function substituteAll(re: RegExp, s: string, replace: (m: RegExpExecArray) => string): { s: string; changed: number } {
  let out = ''
  let last = 0
  let changed = 0
  for (const m of s.matchAll(re)) {
    const replacement = replace(m)
    if (replacement !== m[0]) changed++
    out += s.slice(last, m.index) + replacement
    last = m.index + m[0].length
  }
  return { s: out + s.slice(last), changed }
}

/** fix(), also counting the changes each rule made (curly quotes count in `changes` only). */
export function fixWithRules(
  text: string,
  ctx: WritingContext = {},
  opts: FixOptions = {},
): { text: string; changes: number; changesByRule: Partial<Record<RuleId, number>> } {
  const fragment = opts.fragment ?? false
  const p = protect(text)
  let s = p.s
  let changes = 0
  const changesByRule: Partial<Record<RuleId, number>> = {}
  for (const rule of RULES) {
    const substitute = rule.substitute
    if (!substitute || (rule.fixes && !rule.fixes(ctx))) continue
    const current = s
    const r = substituteAll(rule.re, current, (m) => substitute(m, current, ctx, fragment))
    s = r.s
    changes += r.changed
    if (r.changed) changesByRule[rule.id] = r.changed
  }
  if (opts.quotes) {
    const q = curly(s)
    s = q.text
    changes += q.changes
  }
  return { text: restore(s, p.holes), changes, changesByRule }
}

/**
 * fix(): applies the safe automatic fixes one rule after another and returns the new text and the number of changes.
 * fragment is for a single run inside a paragraph: edge cases that depend on the paragraph boundary are left alone.
 * Throws WritingFixError where the Python raises.
 */
export function fix(text: string, ctx: WritingContext = {}, opts: FixOptions = {}): { text: string; changes: number } {
  const r = fixWithRules(text, ctx, opts)
  return { text: r.text, changes: r.changes }
}

// ------------------------------------------------------------------------------------------- quotes
const DOUBLE_OPENERS = ' \t\n([{\u2014\u2013-/'
const SINGLE_OPENERS = ' \t\n([{\u201C'

function curlyQuote(quote: string, prev: string, next: string): string {
  if (quote === '"') return !prev || DOUBLE_OPENERS.includes(prev) ? '\u201C' : '\u201D'
  return (!prev || SINGLE_OPENERS.includes(prev)) && pyIsAlnum(next) ? '\u2018' : '\u2019'
}

/** curly(): typographic quotes and apostrophes. Assumes URLs and code are already protected. */
export function curly(s: string): { text: string; changes: number } {
  const chars = Array.from(s)
  let text = ''
  let changes = 0
  chars.forEach((ch, i) => {
    if (ch === '"' || ch === "'") {
      text += curlyQuote(ch, chars[i - 1] ?? '', chars[i + 1] ?? '')
      changes++
    } else {
      text += ch
    }
  })
  return { text, changes }
}

/** Straight quotation marks outside protected spans, with the typographic mark curly() would put in their place. */
export function quoteIssues(text: string): WritingIssue[] {
  const p = protect(text)
  const s = p.s
  const out: WritingIssue[] = []
  for (const m of s.matchAll(P.STRAIGHT_QUOTE)) {
    const start = toOriginal(p, m.index, 'start')
    const end = toOriginal(p, m.index + 1, 'end')
    out.push({
      code: 'quotes', severity: 'info', message: 'Straight quotation mark; use a typographic one',
      excerpt: excerptOf(s, m.index, m.index + 1), fixable: true, start, end, match: text.slice(start, end),
      replacement: curlyQuote(m[0], codePointBefore(s, m.index), codePointAt(s, m.index + 1)),
    })
  }
  return out
}
