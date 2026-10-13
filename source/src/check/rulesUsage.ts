// Usage written two ways: contractions in a document otherwise written in full (BwB documents are formal, so the
// contraction is the one to change), and e.g. or i.e. beside "for example" or "that is".

import type { DocSnapshot, Finding } from './types.ts'
import type { IndexedParagraph } from './textIndex.ts'
import { indexText, inQuote } from './textIndex.ts'
import { capped, matchCase, q, replaceFix, spanAt } from './ruleKit.ts'
import { excerpt } from './util.ts'

/** Contractions with one full form. Keys use a straight apostrophe. */
const FULL_FORMS: Readonly<Record<string, string>> = {
  "don't": 'do not', "doesn't": 'does not', "didn't": 'did not', "isn't": 'is not', "aren't": 'are not',
  "wasn't": 'was not', "weren't": 'were not', "hasn't": 'has not', "haven't": 'have not', "hadn't": 'had not',
  "can't": 'cannot', "couldn't": 'could not', "won't": 'will not', "wouldn't": 'would not', "shouldn't": 'should not',
  "mustn't": 'must not', "needn't": 'need not', "we're": 'we are', "they're": 'they are', "you're": 'you are',
  "i'm": 'I am', "we've": 'we have', "they've": 'they have', "you've": 'you have', "i've": 'I have', "we'll": 'we will',
  "they'll": 'they will', "you'll": 'you will', "i'll": 'I will', "it'll": 'it will', "let's": 'let us',
}
/** Contractions that stand for two things (is or has, had or would), which only the author can spell out. */
const AMBIGUOUS = new Set([
  "it's", "that's", "there's", "what's", "here's", "who's", "he's", "she's", "i'd", "we'd", "they'd", "you'd", "it'd",
])
const CONTRACTION = /(?<![\p{L}])\p{L}+['\u2019](?:t|re|ve|ll|m|s|d)(?![\p{L}])/giu
const FULL = /(?<![\p{L}])(?:(?:do|does|did|is|are|was|were|has|have|had|could|would|should|must|need|will) not|cannot|(?:it|that|there|we|they|you) (?:is|are|have|will)|i (?:am|have|will)|let us)(?![\p{L}])/giu

interface Hit {
  ip: IndexedParagraph
  start: number
  end: number
  text: string
}

function hits(paragraphs: IndexedParagraph[], re: RegExp): Hit[] {
  const out: Hit[] = []
  for (const ip of paragraphs) {
    for (const m of ip.p.text.matchAll(re)) {
      if (!inQuote(ip, m.index)) out.push({ ip, start: m.index, end: m.index + m[0].length, text: m[0] })
    }
  }
  return out
}

function contractionFindings(prose: IndexedParagraph[]): Finding[] {
  const found = hits(prose, CONTRACTION).filter((h) => {
    const key = h.text.toLowerCase().replace('\u2019', "'")
    return key in FULL_FORMS || AMBIGUOUS.has(key)
  })
  if (!found.length || !hits(prose, FULL).length) return []
  return capped(found, 30).map((h): Finding => {
    const key = h.text.toLowerCase().replace('\u2019', "'")
    const full = FULL_FORMS[key]
    const base: Finding = {
      id: `consistency:contraction:${h.ip.p.index}:${h.start}`, rule: 'consistency/contractions', group: 'consistency',
      severity: 'copyedit', title: `Contraction ${q(h.text)}`,
      excerpt: excerpt(h.ip.p.text, h.start, h.end), location: spanAt(h.ip.p, h.start, h.end),
    }
    if (full) {
      // "I am" keeps its capital I; other full forms take the case of the contraction.
      const replacement = full.startsWith('I ') ? (h.text === h.text.toUpperCase() ? full.toUpperCase() : full) : matchCase(h.text, full)
      return {
        ...base,
        detail: `The rest of this document is written in full. BwB documents are formal: write ${q(replacement)}.`,
        fix: replaceFix(h.ip.p, h.start, h.end, replacement), fixLabel: `Change to ${q(replacement)}`,
      }
    }
    return {
      ...base, advisory: true,
      detail: key === "it's"
        ? `The rest of this document is written in full. ${q(h.text)} stands for \u201cit is\u201d or \u201cit has\u201d; the possessive is \u201cits\u201d.`
        : 'The rest of this document is written in full. Spell it out: it can stand for is or has, would or had.',
    }
  })
}

/** Each abbreviation with the spelled-out form it is written beside. */
const PAIRS: Array<{ id: string; short: RegExp; long: RegExp; shortName: string; longName: string }> = [
  {
    id: 'eg', shortName: 'e.g.', longName: 'for example',
    short: /(?<![\p{L}])e\.\s?g\.?(?![\p{L}])/gu,
    long: /(?<![\p{L}])for (?:example|instance)(?![\p{L}])/giu,
  },
  {
    id: 'ie', shortName: 'i.e.', longName: 'that is',
    short: /(?<![\p{L}])i\.\s?e\.?(?![\p{L}])/gu,
    // Only "that is" set off as a gloss: ", that is," or "(that is,".
    long: /(?<=[,(\u2013]\s?)that is,/giu,
  },
]

function abbreviationPairFindings(prose: IndexedParagraph[]): Finding[] {
  const out: Finding[] = []
  for (const pair of PAIRS) {
    const short = hits(prose, pair.short)
    const long = hits(prose, pair.long)
    const odd = short.length <= long.length ? short[0] : long[0]
    if (!short.length || !long.length || !odd) continue
    out.push({
      id: `consistency:${pair.id}`, rule: `consistency/${pair.id}`, group: 'consistency', severity: 'copyedit', advisory: true,
      title: `${q(pair.shortName)} (${short.length}) and ${q(pair.longName)} (${long.length}) both used`,
      detail: `Use one throughout. If ${pair.shortName} is kept for brackets and \u201c${pair.longName}\u201d for sentences, that is a consistent rule and this can be left.`,
      excerpt: excerpt(odd.ip.p.text, odd.start, odd.end), location: spanAt(odd.ip.p, odd.start, odd.end),
    })
  }
  return out
}

export function usageFindings(snapshot: DocSnapshot): Finding[] {
  const prose = indexText(snapshot).filter((ip) => ip.kind === 'body' || ip.kind === 'list' || ip.kind === 'table')
  return [...contractionFindings(prose), ...abbreviationPairFindings(prose)]
}
