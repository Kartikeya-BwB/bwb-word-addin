/**
 * BwB headings are in sentence case (titles excepted). A heading reads as title case when, leaving out the first word,
 * small words, all-caps acronyms, words with digits and words straight after a colon, at least three words remain and
 * at least 75 per cent of them are capitalised.
 *
 * The suggestion lower-cases those capitalised words (and capitalised small words), keeps acronyms, words with
 * capitals inside them (iPhone, BwB), single letters (Option A) and the first word and words after a colon. Proper
 * nouns cannot be known, so the suggestion is advice only.
 */
export interface HeadingCaseResult { titleCase: boolean; suggestion?: string }

const SMALL_WORDS = new Set([
  'a', 'an', 'and', 'as', 'at', 'but', 'by', 'for', 'in', 'of', 'on', 'or', 'the', 'to', 'with', 'from', 'into', 'per',
  'via',
])
const HYPHEN = /([-\u2010\u2011])/

interface Word { raw: string; core: string; afterColon: boolean }

/** A word without the punctuation around it: "(And" gives "And", a quoted "Net" gives "Net". */
function coreOf(raw: string): string {
  return raw.replace(/^[^\p{L}\p{N}]+/u, '').replace(/[^\p{L}\p{N}]+$/u, '')
}

const isAcronym = (w: string): boolean => /^\p{Lu}[\p{Lu}\p{N}]*s?$/u.test(w) && (w.match(/\p{Lu}/gu)?.length ?? 0) >= 2
const hasDigit = (w: string): boolean => /\p{N}/u.test(w)
const isCapitalised = (w: string): boolean => /^\p{Lu}/u.test(w)
/** A capital after the first letter of a part, as in iPhone, BwB or McKinsey (Long-Term is two parts). */
const hasInnerCapital = (part: string): boolean => /\p{Lu}/u.test([...part].slice(1).join(''))

function lowerPart(part: string): string {
  const core = coreOf(part)
  if (core.length < 2 || isAcronym(core) || hasDigit(core) || hasInnerCapital(core) || !isCapitalised(core)) return part
  const i = part.indexOf(core)
  return part.slice(0, i) + core.charAt(0).toLowerCase() + core.slice(1) + part.slice(i + core.length)
}

/** Lower-cases a word part by part, keeping the first part when `keepFirst` is set. */
function lowerWord(raw: string, keepFirst: boolean): string {
  return raw.split(HYPHEN).map((part, i) => (HYPHEN.test(part) || (keepFirst && i === 0) ? part : lowerPart(part))).join('')
}

export function checkHeadingCase(text: string): HeadingCaseResult {
  const tokens = text.split(/(\s+)/)
  const words: Word[] = []
  let afterColon = false
  for (const raw of tokens) {
    if (raw === '' || /^\s+$/.test(raw)) continue
    words.push({ raw, core: coreOf(raw), afterColon })
    afterColon = /:[^\p{L}\p{N}]*$/u.test(raw)
  }
  const firstIndex = words.findIndex((w) => /\p{L}/u.test(w.core))
  const eligible = words.filter((w, i) => i > firstIndex && /\p{L}/u.test(w.core) && !w.afterColon &&
    !SMALL_WORDS.has(w.core.toLowerCase()) && !isAcronym(w.core) && !hasDigit(w.core))
  const capitalised = eligible.filter((w) => isCapitalised(w.core)).length
  if (firstIndex < 0 || eligible.length < 3 || capitalised < 0.75 * eligible.length) return { titleCase: false }

  let w = 0
  const suggestion = tokens.map((raw) => {
    if (raw === '' || /^\s+$/.test(raw)) return raw
    const word = words[w]
    const index = w++
    if (!word || index < firstIndex || word.afterColon) return raw
    return lowerWord(raw, index === firstIndex)
  }).join('')
  return { titleCase: true, suggestion }
}
