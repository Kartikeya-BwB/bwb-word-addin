/**
 * The rule table of bwb_lib/writing.py: for each pattern, what scan() reports and what fix() writes. The rules are in
 * the order scan() walks them; fix() applies the ones with a substitution in the same order.
 */
import * as P from './writingPatterns.ts'
import { codePointAt, codePointBefore, cpLength, firstCodePoint, pyInt, pyIsDigit, pyIsUpper, pyRstrip } from './writingUnicode.ts'

export type WritingSeverity = 'warn' | 'info'
export type WritingCode =
  | 'em-dash' | 'en-dash-range' | 'en-dash' | 'en-dash-compound' | 'range-hyphen' | 'per-cent'
  | 'us-spelling' | 'us-spelling-check' | 'date-format' | 'date-numeric' | 'currency' | 'currency-symbol'
  | 'currency-space' | 'brand-name' | 'double-space' | 'placeholder' | 'filler' | 'quotes'
export interface WritingContext { title?: boolean; compact?: boolean }

/** Rule ids are the names of the Python patterns (WARN_* stand for the entries of WARN_ONLY_SPELLING). */
export type RuleId =
  | 'EM_DASH' | 'EN_NUM_RANGE' | 'EN_SPACED' | 'EN_OTHER' | 'YEAR_HYPHEN' | 'PERCENT' | 'OUR_RE' | 'WORDS_RE'
  | 'YZE_RE' | 'ISE_RE' | 'WARN_PROGRAM' | 'WARN_LICENSE' | 'WARN_DIALOG' | 'DATE_US' | 'DATE_ORD' | 'DATE_ABBR_DMY'
  | 'DATE_ABBR_MY' | 'DATE_ISO' | 'DATE_NUMERIC' | 'POUND' | 'EURO' | 'OTHER_SYMBOL' | 'CODE_GLUED' | 'BRAND_BWB'
  | 'BRAND_NAME' | 'DOUBLE_SPACE' | 'PLACEHOLDER' | 'FILLER'

/** The Python exception that fix() raises for the same input. */
export type PythonError = 'KeyError' | 'IndexError' | 'ValueError'

/** Thrown where Python's fix() raises, so callers see the same failures as the Python tools. */
export class WritingFixError extends Error {
  readonly pythonError: PythonError
  constructor(pythonError: PythonError, message: string) {
    super(message)
    this.name = 'WritingFixError'
    this.pythonError = pythonError
  }
}

/** The substitution a rule makes in fix(); `s` is the whole string the rule runs on (Python's m.string). */
export type Substitute = (m: RegExpExecArray, s: string, ctx: WritingContext, fragment: boolean) => string

export interface Rule {
  id: RuleId
  code: WritingCode
  re: RegExp
  severity: (ctx: WritingContext) => WritingSeverity
  message: (match: string) => string
  fixable: (s: string, m: RegExpExecArray, ctx: WritingContext) => boolean
  /** Matches scan() does not report. */
  skip?: (match: string) => boolean
  /** Rules without a substitution are reported only. */
  substitute?: Substitute
  /** fix() leaves abbreviated months alone in table cells. */
  fixes?: (ctx: WritingContext) => boolean
}

// ------------------------------------------------------------------------------------------- helpers
/** _case(): gives the replacement the capitalisation of the source. */
function pyCase(src: string, repl: string): string {
  if (cpLength(src) > 1 && pyIsUpper(src)) return repl.toUpperCase()
  if (pyIsUpper(firstCodePoint(src))) {
    const first = firstCodePoint(repl)
    return first.toUpperCase() + repl.slice(first.length)
  }
  return repl
}

const SENTENCE_SKIP = ' \t\u00A0"\u201C\u2018\'('
const SENTENCE_STOP = '.!?:\n'

function sentenceStart(s: string, i: number): boolean {
  let j = i
  while (j > 0 && SENTENCE_SKIP.includes(s.charAt(j - 1))) j--
  return j === 0 || SENTENCE_STOP.includes(s.charAt(j - 1))
}

/** _proper_noun(): a capitalised match inside a sentence is part of a name (World Health Organization). */
function properNoun(s: string, index: number, tok: string, ctx: WritingContext): boolean {
  if (ctx.title || !pyIsUpper(firstCodePoint(tok)) || (cpLength(tok) > 3 && pyIsUpper(tok))) return false
  return !sentenceStart(s, index)
}

/** _em_sub(): what replaces an em dash and the spaces around it. */
function emDash(s: string, a: number, b: number, match: string, fragment: boolean): string {
  if (a === 0 || b === s.length) return fragment ? match : ''
  if (pyIsDigit(codePointBefore(s, a)) && pyIsDigit(codePointAt(s, b))) return ' to '
  const before = pyRstrip(s.slice(0, a))
  if (before !== '' && '.!?:;,'.includes(before.charAt(before.length - 1))) return ' '
  return ', '
}

const group = (m: RegExpExecArray, i: number): string => m[i] ?? ''
const date = (day: string, month: string, year: string): string => `${pyInt(day)} ${P.monthName(month)} ${year}`

const our: Substitute = (m, s, ctx) => {
  if (properNoun(s, m.index, m[0], ctx)) return m[0]
  const stem = group(m, 1)
  return pyCase(stem, `${stem.slice(0, -2)}our`) + group(m, 2)
}

const word: Substitute = (m, s, ctx) => {
  if (properNoun(s, m.index, m[0], ctx)) return m[0]
  const uk = P.WORDS.get(m[0].toLowerCase())
  if (uk === undefined) throw new WritingFixError('KeyError', `'${m[0].toLowerCase()}'`)
  return pyCase(m[0], uk)
}

const yze: Substitute = (m, s, ctx) => {
  if (properNoun(s, m.index, m[0], ctx)) return m[0]
  const w = m[0]
  const i = group(m, 1).length
  return w.slice(0, i) + (pyIsUpper(w.charAt(i)) ? 'S' : 's') + w.slice(i + 1)
}

const ise: Substitute = (m, s, ctx) => {
  const w = m[0]
  if (P.ISE_EXCEPT.test(w) || properNoun(s, m.index, w, ctx)) return w
  const i = group(m, 1).length + 1 // position of the z
  return w.slice(0, i) + (pyIsUpper(w.charAt(i)) ? 'S' : 's') + w.slice(i + 1)
}

// ------------------------------------------------------------------------------------------- rules
const warn = (): WritingSeverity => 'warn'
const info = (): WritingSeverity => 'info'
const yes = (): boolean => true
const no = (): boolean => false
const say = (message: string) => (): string => message
const usSpelling = (match: string): string => `US spelling '${match}'`
const notProperNoun = (s: string, m: RegExpExecArray, ctx: WritingContext): boolean => !properNoun(s, m.index, m[0], ctx)
const notCompact = (ctx: WritingContext): boolean => !ctx.compact
const abbrFixable = (_s: string, _m: RegExpExecArray, ctx: WritingContext): boolean => !ctx.compact
const abbrSeverity = (ctx: WritingContext): WritingSeverity => (ctx.compact ? 'info' : 'warn')
const ABBREVIATED = 'Abbreviated month; write it in full'

export const RULES: readonly Rule[] = [
  { id: 'EM_DASH', code: 'em-dash', re: P.EM_DASH, severity: warn, fixable: yes,
    message: say('Em dash; use a comma, colon or full stop'),
    substitute: (m, s, _ctx, fragment) => emDash(s, m.index, m.index + m[0].length, m[0], fragment) },
  { id: 'EN_NUM_RANGE', code: 'en-dash-range', re: P.EN_NUM_RANGE, severity: warn, fixable: yes,
    message: say("En dash in a range; write 'to'"), substitute: () => ' to ' },
  { id: 'EN_SPACED', code: 'en-dash', re: P.EN_SPACED, severity: warn, fixable: yes,
    message: say('Spaced en dash; use a comma, colon or full stop'), substitute: () => ', ' },
  { id: 'EN_OTHER', code: 'en-dash-compound', re: P.EN_OTHER, severity: info, fixable: no,
    message: say('En dash between words; use a hyphen or rewrite') },
  { id: 'YEAR_HYPHEN', code: 'range-hyphen', re: P.YEAR_HYPHEN, severity: warn, fixable: yes,
    message: say("Year range with a hyphen; write '2022 to 2025'"), substitute: (m) => `${group(m, 1)} to ${group(m, 2)}` },
  { id: 'PERCENT', code: 'per-cent', re: P.PERCENT, severity: warn, fixable: yes,
    message: say("Write 'per cent' (or use the % sign)"), substitute: (m) => pyCase(m[0], 'per cent') },
  { id: 'OUR_RE', code: 'us-spelling', re: P.OUR_RE, severity: warn, fixable: notProperNoun, message: usSpelling,
    substitute: our },
  { id: 'WORDS_RE', code: 'us-spelling', re: P.WORDS_RE, severity: warn, fixable: notProperNoun, message: usSpelling,
    substitute: word },
  { id: 'YZE_RE', code: 'us-spelling', re: P.YZE_RE, severity: warn, fixable: notProperNoun, message: usSpelling,
    substitute: yze },
  { id: 'ISE_RE', code: 'us-spelling', re: P.ISE_RE, severity: warn, fixable: notProperNoun, message: usSpelling,
    skip: (match) => P.ISE_EXCEPT.test(match), substitute: ise },
  { id: 'WARN_PROGRAM', code: 'us-spelling-check', re: P.WARN_PROGRAM, severity: info, fixable: no,
    message: say("'program' is the US spelling; use 'programme' unless it is a proper noun or software") },
  { id: 'WARN_LICENSE', code: 'us-spelling-check', re: P.WARN_LICENSE, severity: info, fixable: no,
    message: say("'license' is the verb (or US noun); the UK noun is 'licence'") },
  { id: 'WARN_DIALOG', code: 'us-spelling-check', re: P.WARN_DIALOG, severity: info, fixable: no,
    message: say("'dialog' is the US spelling; use 'dialogue' for conversations") },
  { id: 'DATE_US', code: 'date-format', re: P.DATE_US, severity: warn, fixable: yes,
    message: say('Date in US order; write D Month YYYY'), substitute: (m) => date(group(m, 2), group(m, 1), group(m, 3)) },
  { id: 'DATE_ORD', code: 'date-format', re: P.DATE_ORD, severity: warn, fixable: yes,
    message: say('Ordinal in a date; write D Month YYYY'),
    substitute: (m) => `${pyInt(group(m, 1))} ${P.monthName(group(m, 2))}` + (m[3] ? ` ${m[3]}` : '') },
  { id: 'DATE_ABBR_DMY', code: 'date-format', re: P.DATE_ABBR_DMY, severity: abbrSeverity, fixable: abbrFixable,
    fixes: notCompact, message: say(ABBREVIATED), substitute: (m) => date(group(m, 1), group(m, 2), group(m, 3)) },
  { id: 'DATE_ABBR_MY', code: 'date-format', re: P.DATE_ABBR_MY, severity: abbrSeverity, fixable: abbrFixable,
    fixes: notCompact, message: say(ABBREVIATED), substitute: (m) => `${P.monthName(group(m, 1))} ${group(m, 2)}` },
  { id: 'DATE_ISO', code: 'date-format', re: P.DATE_ISO, severity: warn, fixable: yes,
    message: say('ISO date in prose; write D Month YYYY'),
    substitute: (m) => date(group(m, 3), P.MONTHS[Number(group(m, 2)) - 1] ?? '', group(m, 1)) },
  { id: 'DATE_NUMERIC', code: 'date-numeric', re: P.DATE_NUMERIC, severity: info, fixable: no,
    message: say('Numeric date; write the month in words (D Month YYYY)') },
  { id: 'POUND', code: 'currency', re: P.POUND, severity: warn, fixable: yes,
    message: say('Use the ISO code GBP instead of the pound sign'), substitute: () => 'GBP ' },
  { id: 'EURO', code: 'currency', re: P.EURO, severity: warn, fixable: yes,
    message: say('Use the ISO code EUR instead of the euro sign'), substitute: () => 'EUR ' },
  { id: 'OTHER_SYMBOL', code: 'currency-symbol', re: P.OTHER_SYMBOL, severity: warn, fixable: no,
    message: say('Use an ISO currency code (USD, CAD, AUD ...); check which currency is meant') },
  { id: 'CODE_GLUED', code: 'currency-space', re: P.CODE_GLUED, severity: info, fixable: yes,
    message: say('Put a space between the currency code and the figure'), substitute: (m) => `${m[0]} ` },
  { id: 'BRAND_BWB', code: 'brand-name', re: P.BRAND_BWB, severity: warn, fixable: yes,
    message: say("Write the abbreviation as 'BwB'"), substitute: () => 'BwB' },
  { id: 'BRAND_NAME', code: 'brand-name', re: P.BRAND_NAME, severity: warn, fixable: yes,
    message: say(`Write '${P.BRAND_NAME_TEXT}'`), skip: (match) => match === P.BRAND_NAME_TEXT,
    substitute: () => P.BRAND_NAME_TEXT },
  { id: 'DOUBLE_SPACE', code: 'double-space', re: P.DOUBLE_SPACE, severity: info, fixable: yes,
    message: say('Double space'), substitute: () => ' ' },
  { id: 'PLACEHOLDER', code: 'placeholder', re: P.PLACEHOLDER, severity: warn, fixable: no,
    message: (match) => `Placeholder text left in: '${match}'` },
  { id: 'FILLER', code: 'filler', re: P.FILLER, severity: info, fixable: no, message: (match) => `Filler phrase '${match}'` },
]
