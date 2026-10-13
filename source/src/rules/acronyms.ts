/**
 * Acronyms in a document: where each is first used and whether, and where, it is spelled out.
 *
 * Candidates are tokens of 2 to 8 ASCII letters and digits that start with a capital and have at least two capitals,
 * no more lower-case letters than capitals (NbS, AfDB, KfW) and an optional plural "s" (SDGs, MDBs); a capital followed
 * by two or more digits (G20) also counts. Not candidates: Roman numerals, ISO 4217 currency codes, deliverable and
 * section codes such as D7.4, anything inside URLs, e-mail addresses, code spans and icon tokens, and words in an
 * all-caps phrase (EXECUTIVE SUMMARY), which is emphasis rather than a list of acronyms.
 *
 * Four currency codes are far more often acronyms in development finance (SDG, COP, IRR, SSP); they are skipped only
 * when an amount follows (SDG 5m, COP 5,000), not a goal or session number (SDG 7, COP 28). Chemical formulas (CO2,
 * CH4) are not acronyms either.
 *
 * Definitions are "Full Name (ACR)" and "ACR (Full Name)", where the initials of the words match the letters of the
 * acronym case-insensitively. Joining words (and, of, for, the ...) may give a letter or none (Conference of the
 * Parties, COP), each part of a hyphenated word may give its initial (Nature-based Solutions, NbS), and a lower-case
 * letter of the acronym may also come from inside the word before it (African Development Bank, AfDB). A row of a
 * two-column abbreviations table (acronym cell, then expansion cell) is a definition too. Headings count as uses but
 * never as definitions.
 */
import { protectPattern } from './writingPatterns.ts'

export interface AcronymInputParagraph { text: string; isHeading?: boolean; inTable?: boolean }
export type AcronymStatus = 'defined' | 'defined-late' | 'undefined' | 'common'
export interface AcronymUse { paragraph: number; start: number; end: number }
export interface AcronymEntry {
  acronym: string
  /** Expansion found in the document, if any. */
  definition?: string
  /** Where the definition pattern was found. */
  definedAt?: AcronymUse
  firstUse: AcronymUse
  uses: number
  status: AcronymStatus
}
/** One place an acronym is spelled out: in the text, or as a row of a two-column abbreviations table. */
export interface AcronymDefinition { definition: string; at: AcronymUse; table: boolean }
/** The entries, with every use and every definition of each acronym, keyed by the acronym as written. */
export interface AcronymAnalysis {
  entries: AcronymEntry[]
  uses: ReadonlyMap<string, readonly AcronymUse[]>
  definitions: ReadonlyMap<string, readonly AcronymDefinition[]>
}

const COMMON = new Set(['UK', 'US', 'EU', 'UN', 'BWB']) // compared upper-cased
// Chemical formulas that look like acronyms in climate texts.
const FORMULAS = new Set(['CO2', 'CO2E', 'CH4', 'N2O', 'NO2', 'NOX', 'SO2', 'SOX', 'NH3', 'SF6', 'H2O', 'O3'])
const JOINING_WORDS = new Set(['and', 'of', 'for', 'the', 'in', 'on', 'to', 'de', 'la', 'et', '&'])
const ACRONYM_FIRST = new Set(['COP', 'IRR', 'SDG', 'SSP'])
// ISO 4217: active codes, fund codes, precious metals and recently withdrawn codes still seen in documents.
const ISO_4217 = new Set((
  'AED AFN ALL AMD ANG AOA ARS AUD AWG AZN BAM BBD BDT BGN BHD BIF BMD BND BOB BOV BRL BSD BTN BWP BYN BYR BZD CAD ' +
  'CDF CHE CHF CHW CLF CLP CNH CNY COP COU CRC CUC CUP CVE CYP CZK DEM DJF DKK DOP DZD EEK EGP ERN ESP ETB EUR FIM ' +
  'FJD FKP FRF GBP GEL GHS GIP GMD GNF GRD GTQ GYD HKD HNL HRK HTG HUF IDR IEP ILS INR IQD IRR ISK ITL JMD JOD JPY ' +
  'KES KGS KHR KMF KPW KRW KWD KYD KZT LAK LBP LKR LRD LSL LTL LVL LYD MAD MDL MGA MKD MMK MNT MOP MRO MRU MTL MUR ' +
  'MVR MWK MXN MXV MYR MZN NAD NGN NIO NLG NOK NPR NZD OMR PAB PEN PGK PHP PKR PLN PTE PYG QAR ROL RON RSD RUB RWF ' +
  'SAR SBD SCR SDG SEK SGD SHP SIT SKK SLE SLL SOS SRD SSP STD STN SVC SYP SZL THB TJS TMT TND TOP TRY TTD TWD TZS ' +
  'UAH UGX USD USN UYI UYU UYW UZS VEB VED VEF VES VND VUV WST XAF XAG XAU XBA XBB XBC XBD XCD XCG XDR XOF XPD XPF ' +
  'XPT XSU XTS XUA XXX YER ZAR ZMK ZMW ZWD ZWG ZWL'
).split(' '))
const ROMAN = /^(?=[IVXL]+$)(?:XL|L?X{0,3})(?:IX|IV|V?I{0,3})$/
const TOKEN = /[\p{L}\p{N}]+/gu
const HYPHEN = /[-\u2010\u2011]/
const SENTENCE_END = /[.;:!?]$/

interface Token { text: string; start: number; end: number }
interface Candidate { key: string; start: number; end: number }
interface Definition { key: string; definition: string; at: AcronymUse; table?: boolean }

// ------------------------------------------------------------------------------------------- candidates
function shape(s: string): boolean {
  if (!/^[A-Z][A-Za-z0-9]*$/.test(s)) return false
  const upper = s.replace(/[^A-Z]/g, '').length
  const lower = s.replace(/[^a-z]/g, '').length
  return (upper >= 2 && upper >= lower) || /^[A-Z][0-9]{2,}$/.test(s)
}

/** The acronym a token stands for (its singular), or null when the token is not an acronym. */
function acronymOf(token: string): string | null {
  if (token.length < 2 || token.length > 8) return null
  if (/[A-Z0-9]s$/.test(token) && shape(token.slice(0, -1))) return token.slice(0, -1)
  return shape(token) ? token : null
}

// An amount after a currency code: 5m, 1.2bn, 300 million, 5,000.
const AMOUNT_AFTER = /^\s?(?:\d[\d,]*(?:\.\d+)?\s?(?:m|bn|k|tn|million|billion|thousand|trillion)\b|\d{1,3}(?:,\d{3})+)/i

function excluded(key: string, text: string, t: Token): boolean {
  if (ROMAN.test(key)) return true
  if (FORMULAS.has(key.toUpperCase())) return true
  if (ISO_4217.has(key) && (!ACRONYM_FIRST.has(key) || AMOUNT_AFTER.test(text.slice(t.end, t.end + 24)))) return true
  // Deliverable, task and section codes: D7.4, WP3.1, 2.D1
  return /^\.\d/.test(text.slice(t.end, t.end + 2)) || /\d\.$/.test(text.slice(Math.max(0, t.start - 2), t.start))
}

const isCapsWord = (w: string): boolean => /\p{L}/u.test(w) && !/\p{Ll}/u.test(w)

/** Tokens inside an all-caps phrase: three or more capitalised words in a row, or two with one of five letters. */
function capsPhraseTokens(text: string, tokens: readonly Token[]): Set<number> {
  const out = new Set<number>()
  let run: number[] = []
  const flush = (): void => {
    const long = run.some((i) => (tokens[i]?.text.replace(/[^\p{L}]/gu, '').length ?? 0) >= 5)
    if (run.length >= 3 || (run.length === 2 && long)) for (const i of run) out.add(i)
    run = []
  }
  tokens.forEach((t, i) => {
    const prev = tokens[i - 1]
    const joined = prev !== undefined && /^\s+$/.test(text.slice(prev.end, t.start))
    if (!isCapsWord(t.text) || (run.length > 0 && !joined)) flush()
    if (isCapsWord(t.text)) run.push(i)
  })
  flush()
  return out
}

function candidates(text: string): Candidate[] {
  const masked = text.replace(protectPattern(text), (m) => ' '.repeat(m.length))
  const tokens: Token[] = [...masked.matchAll(TOKEN)]
    .map((m) => ({ text: m[0], start: m.index, end: m.index + m[0].length }))
  const caps = capsPhraseTokens(masked, tokens)
  const out: Candidate[] = []
  tokens.forEach((t, i) => {
    const key = acronymOf(t.text)
    if (key && !caps.has(i) && !excluded(key, masked, t)) out.push({ key, start: t.start, end: t.end })
  })
  return out
}

// ------------------------------------------------------------------------------------------- definitions
interface Letter { ch: string; lower: boolean }
interface Word { clean: string; start: number; end: number }

function lettersOf(acronym: string): Letter[] {
  return [...acronym]
    .filter((c) => /[A-Za-z]/.test(c))
    .map((c) => ({ ch: c.toLowerCase(), lower: c !== c.toUpperCase() }))
}

/** Letter positions reachable after one word gives its initial, and optionally more (see the module comment). */
function afterWord(word: string, letters: readonly Letter[], from: number): number[] {
  const parts = word.split(HYPHEN).filter((p) => p !== '')
  const out = new Set<number>()
  const visit = (p: number, li: number, required: boolean): void => {
    const part = parts[p]
    if (part === undefined) {
      out.add(li)
      return
    }
    if (!required) visit(p + 1, li, false)
    if (part.charAt(0).toLowerCase() !== letters[li]?.ch) return
    let l = li + 1
    visit(p + 1, l, false)
    for (let k = 1; k < part.length && letters[l]?.lower && part.charAt(k).toLowerCase() === letters[l]?.ch; k++) {
      l++
      visit(p + 1, l, false)
    }
  }
  visit(0, from, true)
  return [...out]
}

/** Do all these words, in order, spell the letters? The first word must give a letter. */
function spells(words: readonly Word[], letters: readonly Letter[]): boolean {
  if (letters.length === 0) return false
  const memo = new Map<number, boolean>()
  const go = (wi: number, li: number): boolean => {
    const word = words[wi]
    if (word === undefined) return li === letters.length
    const key = wi * 64 + li
    const known = memo.get(key)
    if (known !== undefined) return known
    const skippable = wi > 0 && (JOINING_WORDS.has(word.clean.toLowerCase()) || !/\p{L}/u.test(word.clean))
    const ok = (skippable && go(wi + 1, li)) || afterWord(word.clean, letters, li).some((next) => go(wi + 1, next))
    memo.set(key, ok)
    return ok
  }
  return go(0, 0)
}

function wordsOf(text: string, offset: number): Word[] {
  const out: Word[] = []
  for (const m of text.matchAll(/\S+/gu)) {
    const lead = /^[^\p{L}\p{N}&]*/u.exec(m[0])?.[0].length ?? 0
    const clean = m[0].slice(lead).replace(/[^\p{L}\p{N}]+$/u, '')
    if (clean !== '') out.push({ clean, start: offset + m.index + lead, end: offset + m.index + lead + clean.length })
  }
  return out
}

function definitions(text: string, paragraph: number): Definition[] {
  const masked = text.replace(protectPattern(text), (m) => ' '.repeat(m.length))
  const out: Definition[] = []
  // Full Name (ACR)
  for (const m of masked.matchAll(/\(\s*([A-Za-z0-9]+)\s*\)/g)) {
    const key = acronymOf(m[1] ?? '')
    if (!key) continue
    const letters = lettersOf(key)
    const before = masked.slice(0, m.index)
    const all = wordsOf(before, 0)
    // The expansion cannot reach back past the end of the previous sentence.
    const lastSentenceEnd = all.findLastIndex((w, i) => {
      const next = all[i + 1]
      return next !== undefined && SENTENCE_END.test(before.slice(w.start, next.start).trim())
    })
    const words = all.slice(lastSentenceEnd + 1).slice(-(letters.length * 2 + 2))
    // The shortest run of words that ends at the parenthesis and spells the acronym.
    for (let j = words.length - 1; j >= 0; j--) {
      const span = words.slice(j)
      const first = span[0]
      const last = span[span.length - 1]
      if (first && last && spells(span, letters)) {
        const at = { paragraph, start: first.start, end: m.index + m[0].length }
        out.push({ key, definition: text.slice(first.start, last.end), at })
        break
      }
    }
  }
  // ACR (Full Name)
  for (const m of masked.matchAll(/(?<![\p{L}\p{N}])([A-Za-z0-9]+)\s*\(([^()]+)\)/gu)) {
    const key = acronymOf(m[1] ?? '')
    const inner = m[2] ?? ''
    if (!key) continue
    const words = wordsOf(inner, m.index + m[0].length - 1 - inner.length)
    const letters = lettersOf(key)
    // All the words in the parentheses, after any leading joining words ("the").
    for (let i = 0; i < words.length; i++) {
      const span = words.slice(i)
      const first = span[0]
      const last = span[span.length - 1]
      if (first && last && spells(span, letters)) {
        const at = { paragraph, start: m.index, end: m.index + m[0].length }
        out.push({ key, definition: text.slice(first.start, last.end), at })
        break
      }
      if (!JOINING_WORDS.has(first?.clean.toLowerCase() ?? '')) break
    }
  }
  return out.sort((a, b) => a.at.start - b.at.start)
}

/** A two-column abbreviations table: a cell holding only the acronym, then a cell holding its expansion. */
function tableRowDefinition(
  cell: AcronymInputParagraph,
  next: AcronymInputParagraph | undefined,
  paragraph: number,
): Definition | null {
  if (!cell.inTable || !next?.inTable || next.isHeading) return null
  const token = cell.text.trim()
  const key = acronymOf(token)
  if (!key || !/^[A-Za-z0-9]+$/.test(token)) return null
  const words = wordsOf(next.text, 0)
  const first = words[0]
  const last = words[words.length - 1]
  if (!first || !last || !spells(words, lettersOf(key))) return null
  const start = cell.text.indexOf(token)
  return { key, definition: next.text.slice(first.start, last.end), at: { paragraph, start, end: start + token.length }, table: true }
}

// ------------------------------------------------------------------------------------------- analysis
const compare = (a: string, b: string): number => (a < b ? -1 : a > b ? 1 : 0)
const before = (a: AcronymUse, b: AcronymUse): boolean =>
  a.paragraph < b.paragraph || (a.paragraph === b.paragraph && a.end <= b.start)
const inside = (a: AcronymUse, b: AcronymUse): boolean =>
  a.paragraph === b.paragraph && a.start >= b.start && a.end <= b.end

/** Every acronym in the document, sorted alphabetically. */
export function analyseAcronyms(paragraphs: AcronymInputParagraph[]): AcronymEntry[] {
  return acronymAnalysis(paragraphs).entries.slice()
}

let memo: { input: AcronymInputParagraph[]; result: AcronymAnalysis } | null = null

const sameInput = (a: readonly AcronymInputParagraph[], b: readonly AcronymInputParagraph[]): boolean =>
  a.length === b.length && a.every((p, i) => {
    const q = b[i]
    return q !== undefined && p.text === q.text && !p.isHeading === !q.isHeading && !p.inTable === !q.inTable
  })

/**
 * The analysis behind analyseAcronyms, with every use and definition. The last result is kept, so checks that ask
 * about the same paragraphs one after another share one pass over the document.
 */
export function acronymAnalysis(paragraphs: AcronymInputParagraph[]): AcronymAnalysis {
  if (memo && sameInput(memo.input, paragraphs)) return memo.result
  const uses = new Map<string, AcronymUse[]>()
  const defs = new Map<string, Definition>()
  const allDefs = new Map<string, AcronymDefinition[]>()
  paragraphs.forEach((p, i) => {
    for (const c of candidates(p.text)) {
      const list = uses.get(c.key) ?? []
      list.push({ paragraph: i, start: c.start, end: c.end })
      uses.set(c.key, list)
    }
    if (p.isHeading) return
    const found = [tableRowDefinition(p, paragraphs[i + 1], i), ...definitions(p.text, i)]
    for (const d of found) {
      if (!d) continue
      if (!defs.has(d.key)) defs.set(d.key, d)
      allDefs.set(d.key, [...(allDefs.get(d.key) ?? []), { definition: d.definition, at: d.at, table: d.table === true }])
    }
  })
  const result: AcronymAnalysis = { entries: entriesOf(uses, defs), uses, definitions: allDefs }
  memo = { input: paragraphs.map((p) => ({ ...p })), result }
  return result
}

function entriesOf(uses: ReadonlyMap<string, AcronymUse[]>, defs: ReadonlyMap<string, Definition>): AcronymEntry[] {
  const out: AcronymEntry[] = []
  for (const [acronym, list] of uses) {
    const firstUse = list[0]
    if (!firstUse) continue
    const def = defs.get(acronym)
    const entry: AcronymEntry = { acronym, firstUse, uses: list.length, status: 'undefined' }
    if (def) {
      entry.definition = def.definition
      entry.definedAt = def.at
    }
    if (COMMON.has(acronym.toUpperCase())) entry.status = 'common'
    else if (def) entry.status = list.some((u) => before(u, def.at) && !inside(u, def.at)) ? 'defined-late' : 'defined'
    out.push(entry)
  }
  return out.sort((a, b) => compare(a.acronym.toUpperCase(), b.acronym.toUpperCase()) || compare(a.acronym, b.acronym))
}
