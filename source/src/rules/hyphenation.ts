/**
 * Compounds written more than one way in the same document: hyphenated (decision-making), open (decision making) and
 * closed (decisionmaking), compared case-insensitively, in the spirit of PerfectIt's hyphenation check.
 *
 * Only two-part compounds of letters are considered, so numbers and ranges (2022-2025, 5-10, COVID-19) and longer chains
 * (state-of-the-art) never count. An open form counts only when its two words also appear hyphenated or closed, and
 * not when the first word is a function word (may be, in to, any one) or the second a particle (set up, follow up,
 * roll out): there the open form is usually a verb and the other form a noun, and both are right.
 */
export interface HyphenVariant { forms: Array<{ form: string; count: number }>; suggestion: string }

type Kind = 'hyphenated' | 'open' | 'closed'
interface Tally { count: number; written: string; first: number }

const WORD = /[\p{L}\p{M}]+/gu
const HYPHENS = new Set(['-', '\u2010', '\u2011'])
const OPEN_GAPS = new Set([' ', '\u00A0'])
const FUNCTION_WORDS = new Set(['a', 'an', 'the', 'in', 'on', 'to', 'at', 'by', 'of', 'may', 'any', 'every', 'some', 'all', 'no'])
const PARTICLES = new Set(['up', 'down', 'out', 'off', 'in', 'on', 'over', 'back', 'through', 'away', 'about', 'around'])
// Compounds that are mostly adjectives before a noun (a long-term plan): on a tie the hyphenated form wins.
const ADJECTIVE_FIRST = new Set([
  'long', 'short', 'medium', 'mid', 'high', 'low', 'full', 'part', 'large', 'small', 'well', 'cross', 'middle', 'upper',
  'lower', 'multi', 'single', 'half', 'first', 'last', 'top', 'open', 'fast', 'slow', 'real', 'world',
])
const ADJECTIVE_SECOND = new Set([
  'term', 'scale', 'level', 'income', 'time', 'based', 'related', 'led', 'driven', 'wide', 'oriented', 'specific',
  'sensitive', 'friendly', 'free', 'intensive', 'owned', 'sized', 'making', 'focused', 'proof', 'neutral', 'facing',
  'funded', 'backed', 'linked', 'aligned', 'compliant', 'ready', 'efficient', 'class', 'quality', 'risk', 'cost',
])

function tally(map: Map<string, Tally>, key: string, written: string, first: number): void {
  const t = map.get(key)
  if (t) t.count++
  else map.set(key, { count: 1, written, first })
}

/** Compounds that appear in more than one form, in order of first appearance. */
export function findHyphenationVariants(texts: string[]): HyphenVariant[] {
  const forms: Record<Kind, Map<string, Tally>> = { hyphenated: new Map(), open: new Map(), closed: new Map() }
  let position = 0
  for (const text of texts) {
    const words = [...text.matchAll(WORD)].map((m) => ({ lower: m[0].toLowerCase(), start: m.index, end: m.index + m[0].length }))
    const gap = (i: number): string => {
      const a = words[i]
      const b = words[i + 1]
      return a && b ? text.slice(a.end, b.start) : ''
    }
    words.forEach((w, i) => {
      const next = words[i + 1]
      tally(forms.closed, w.lower, w.lower, position++)
      if (!next) return
      const g = gap(i)
      const key = `${w.lower} ${next.lower}`
      if (HYPHENS.has(g)) {
        const chained = HYPHENS.has(gap(i - 1)) || HYPHENS.has(gap(i + 1))
        if (!chained && w.lower.length + next.lower.length > 2) tally(forms.hyphenated, key, `${w.lower}${g}${next.lower}`, position)
      } else if (OPEN_GAPS.has(g) && !FUNCTION_WORDS.has(w.lower) && !PARTICLES.has(next.lower)) {
        tally(forms.open, key, `${w.lower}${g}${next.lower}`, position)
      }
    })
  }

  const pairs = new Set(forms.hyphenated.keys())
  for (const key of forms.open.keys()) if (forms.closed.has(key.replace(' ', ''))) pairs.add(key)
  const out: Array<HyphenVariant & { first: number }> = []
  for (const key of pairs) {
    const [w1 = '', w2 = ''] = key.split(' ')
    const adjective = ADJECTIVE_FIRST.has(w1) || ADJECTIVE_SECOND.has(w2)
    const preference: Kind[] = adjective ? ['hyphenated', 'closed', 'open'] : ['closed', 'hyphenated', 'open']
    const found = preference
      .map((kind) => ({ kind, t: forms[kind].get(kind === 'closed' ? w1 + w2 : key) }))
      .filter((x): x is { kind: Kind; t: Tally } => x.t !== undefined)
    if (found.length < 2) continue
    // Stable sort: equal counts keep the preference order, so the first entry is the suggestion.
    found.sort((a, b) => b.t.count - a.t.count)
    out.push({
      forms: found.map((x) => ({ form: x.t.written, count: x.t.count })),
      suggestion: found[0]?.t.written ?? '',
      first: Math.min(...found.map((x) => x.t.first)),
    })
  }
  return out.sort((a, b) => a.first - b.first).map(({ forms: f, suggestion }) => ({ forms: f, suggestion }))
}
