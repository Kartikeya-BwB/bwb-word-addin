/**
 * Changing case the BwB way. Acronyms (two or more capitals: UNDP, BwB, MDBs) and names with a capital inside (iPhone,
 * McKinsey) keep their capitals in sentence, title and lower case. Title case leaves small words (and, of, the ...) in
 * lower case unless they start the text, end it or follow a colon. Sentence case capitalises the start of the text and
 * of each sentence, and the pronoun I.
 *
 * Text written all in capitals (EXECUTIVE SUMMARY) cannot show which words are acronyms, so there only the acronyms
 * the caller knows (those used elsewhere in the document), a few common ones and words with digits (COP28) are kept,
 * and short words with no vowel (GCF, MDB) are taken for acronyms. BWB becomes BwB.
 */

export type CaseMode = 'sentence' | 'title' | 'upper' | 'lower'

export interface CaseOptions {
  /** Acronyms to keep in text written all in capitals. */
  acronyms?: ReadonlySet<string>
}

const SMALL_WORDS = new Set([
  'a', 'an', 'and', 'as', 'at', 'but', 'by', 'for', 'from', 'in', 'into', 'nor', 'of', 'on', 'onto', 'or', 'per', 'so',
  'the', 'to', 'up', 'via', 'vs', 'with', 'yet',
])
/** Abbreviations after which a full stop does not end a sentence (compared in lower case). */
const ABBREVIATIONS = new Set([
  'e.g.', 'i.e.', 'etc.', 'cf.', 'vs.', 'approx.', 'no.', 'nos.', 'fig.', 'p.', 'pp.', 'vol.', 'ed.', 'eds.', 'mr.', 'mrs.',
  'ms.', 'dr.', 'st.',
])
const COMMON_ACRONYMS = new Set([
  'UK', 'US', 'USA', 'EU', 'UN', 'UNDP', 'UNEP', 'UNFCCC', 'OECD', 'IMF', 'WHO', 'IFC', 'EIB', 'EBRD', 'AFDB', 'ADB', 'GCF',
  'GEF', 'MDB', 'DFI', 'SDG', 'COP', 'ESG', 'NDC', 'IPCC', 'SME', 'MSME', 'NGO', 'CEO', 'CFO', 'GDP', 'FAO', 'IFAD',
])
/** Words BwB always writes one way, whatever the case of the text around them. */
const FIXED: Readonly<Record<string, string>> = { BWB: 'BwB', BWBS: 'BwBs' }

const LETTER = /\p{L}/u
const HYPHEN = /([-\u2010\u2011])/
const CLOSERS = /["'\u2019\u201d)\]]*$/
const capitals = (w: string): number => w.match(/\p{Lu}/gu)?.length ?? 0
const hasLower = (w: string): boolean => /\p{Ll}/u.test(w)
const hasDigit = (w: string): boolean => /\p{N}/u.test(w)

/** Splits a word into the punctuation before it, its core and the punctuation after it: "(UNDP)," gives "(", "UNDP", "),". */
function partsOf(raw: string): [string, string, string] {
  const lead = /^[^\p{L}\p{N}]*/u.exec(raw)?.[0] ?? ''
  const rest = raw.slice(lead.length)
  const trail = /[^\p{L}\p{N}]*$/u.exec(rest)?.[0] ?? ''
  return [lead, rest.slice(0, rest.length - trail.length), trail]
}

/** Text in capitals throughout: two or more words, or one long word, with no lower-case letter. */
export function isShouting(text: string): boolean {
  const words = text.split(/\s+/).map((w) => partsOf(w)[1]).filter((w) => (w.match(/\p{L}/gu)?.length ?? 0) >= 2)
  if (!words.length || words.some(hasLower)) return false
  return words.length >= 2 || (words[0]?.length ?? 0) >= 6
}

interface Context {
  shouting: boolean
  acronyms: ReadonlySet<string>
}

/** True when a word part keeps its capitals. */
function keeps(part: string, c: Context): boolean {
  if (!LETTER.test(part)) return true
  const base = part.replace(/['\u2019][sS]$/u, '')
  if (c.shouting) {
    const singular = /^\p{Lu}{2,}S$/u.test(base) ? base.slice(0, -1) : base
    return [base, singular].some((w) => c.acronyms.has(w) || COMMON_ACRONYMS.has(w)) || hasDigit(base) ||
      (/^\p{Lu}{2,5}$/u.test(base) && !/[AEIOUY]/.test(base))
  }
  return capitals(base) >= 2 || /^.+\p{Lu}/u.test(base) || (capitals(base) === 1 && hasDigit(base))
}

const upperFirst = (w: string): string => {
  const i = w.search(LETTER)
  return i < 0 ? w : w.slice(0, i) + w.charAt(i).toUpperCase() + w.slice(i + 1)
}

/** Applies `change` to each hyphen-separated part of a word. */
function byPart(core: string, change: (part: string, index: number) => string): string {
  let index = 0
  return core.split(HYPHEN).map((p) => (HYPHEN.test(p) ? p : change(p, index++))).join('')
}

/** The company name and the like, written their one way; null for other words. */
const fixed = (part: string, c: Context): string | null => (c.shouting ? (FIXED[part.toUpperCase()] ?? null) : null)

const lowerCore = (core: string, c: Context): string =>
  byPart(core, (p) => fixed(p, c) ?? (keeps(p, c) ? p : p.toLowerCase()))

const PRONOUN = /^i(?:['\u2019](?:m|ve|d|ll))?$/i

/** Changes the case of text, keeping its spaces and punctuation. */
export function changeCase(text: string, mode: CaseMode, options: CaseOptions = {}): string {
  if (mode === 'upper') return text.toUpperCase()
  const c: Context = { shouting: isShouting(text), acronyms: options.acronyms ?? new Set() }
  const tokens = text.split(/(\s+)/)
  let lastWord = -1
  tokens.forEach((t, i) => {
    if (partsOf(t)[1]) lastWord = i
  })
  let sentenceStart = true
  let afterColon = false
  let first = true
  return tokens.map((raw, i) => {
    if (!raw || /^\s+$/.test(raw)) return raw
    const [lead, core, trail] = partsOf(raw)
    let out = core
    if (mode === 'lower') out = lowerCore(core, c)
    else if (mode === 'sentence') {
      out = lowerCore(core, c)
      // A sentence starts with a capital, unless its first word keeps its own (iPhone, UN-backed).
      if (PRONOUN.test(core) || (sentenceStart && core && !keeps(core.split(HYPHEN)[0] ?? core, c))) out = upperFirst(out)
    } else {
      out = byPart(core, (p, part) => {
        const known = fixed(p, c)
        if (known) return known
        if (keeps(p, c)) return p
        const edge = part === 0 && (first || i === lastWord || afterColon)
        return SMALL_WORDS.has(p.toLowerCase()) && !edge ? p.toLowerCase() : upperFirst(p.toLowerCase())
      })
    }
    if (core) {
      first = false
      const word = `${core}${trail.replace(CLOSERS, '')}`.toLowerCase()
      sentenceStart = /[.!?]["'\u2019\u201d)\]]*$/.test(trail) && !ABBREVIATIONS.has(word)
      afterColon = /:["'\u2019\u201d)\]]*$/.test(trail)
    }
    return lead + out + trail
  }).join('')
}
