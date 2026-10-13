// Terms written more than one way: the same words with and without capitals mid-sentence (Green Bonds and green bonds),
// compound words in forms the hyphenation check does not pair because they differ in number (policymakers and policy
// maker), and accents (Cote d'Ivoire with and without its circumflex). All are suggestions: only the author knows
// which form is meant.

import type { DocSnapshot, Finding } from './types.ts'
import type { IndexedParagraph, Word } from './textIndex.ts'
import { atSentenceStart, indexText, inQuote, joinedBySpace, wordsOf } from './textIndex.ts'
import { capped, matchCase, q, replaceFix, spanAt } from './ruleKit.ts'
import { excerpt } from './util.ts'

interface Occurrence {
  ip: IndexedParagraph
  start: number
  end: number
}

/** Small words that may sit inside a name: Bankers without Boundaries, Ministry of Finance. */
const JOINERS = new Set(['of', 'the', 'for', 'and', 'on', 'in', 'to', 'at', 'de', 'la', 'du', 'des', 'del', 'without'])
/** Words a sentence often starts with, which say nothing when capitalised. */
const FUNCTION_WORDS = new Set((
  'a an the this that these those it its we our us they their them he she his her you your i in on at by for from with ' +
  'without as but and or nor if when while where which who whom whose what how why all each every some many most more ' +
  'much no not any other another both either neither such there here however also after before since during under over ' +
  'between through into onto about above below than then thus so yet still first second third finally overall one'
).split(' '))
const NAMES_OF_TIME = new Set((
  'january february march april may june july august september october november december monday tuesday wednesday ' +
  'thursday friday saturday sunday'
).split(' '))

const isCapitalised = (w: string): boolean =>
  /^\p{Lu}/u.test(w) && /\p{Ll}/u.test(w) && w.split(/[-\u2010\u2011]/).every((part) => !/\p{Lu}/u.test(part.slice(1)))
const isLowerCase = (w: string): boolean => /\p{Ll}/u.test(w) && !/\p{Lu}/u.test(w)
const runText = (ip: IndexedParagraph, o: { start: number; end: number }): string => ip.p.text.slice(o.start, o.end)
const keyOf = (s: string): string => s.toLowerCase().replace(/\s+/g, ' ')

/** Runs of capitalised words mid-sentence, up to four of them, that are not part of a longer name or a label. */
function capitalRuns(ip: IndexedParagraph, words: Word[]): Occurrence[] {
  const text = ip.p.text
  const out: Occurrence[] = []
  let i = 0
  while (i < words.length) {
    let first = words[i]
    if (!first || !isCapitalised(first.text)) {
      i += 1
      continue
    }
    let j = i
    let caps = 1
    for (;;) {
      const a = words[j]
      const b = words[j + 1]
      const c = words[j + 2]
      if (!a || !b || !joinedBySpace(text, a, b)) break
      if (isCapitalised(b.text)) {
        j += 1
        caps += 1
      } else if (JOINERS.has(b.lower) && c && joinedBySpace(text, b, c) && isCapitalised(c.text)) {
        j += 2
        caps += 1
      } else break
    }
    const before = words[i - 1]
    const last = words[j] ?? first
    const start = i
    i = j + 1
    // "The Board": the capital of a sentence's first small word says nothing; the rest of the run still counts.
    if (atSentenceStart(ip, first.start) && FUNCTION_WORDS.has(first.lower) && caps > 1) {
      first = words[start + 1] ?? first
      caps -= 1
    }
    if (caps > 4 || atSentenceStart(ip, first.start) || inQuote(ip, first.start)) continue
    if (caps === 1 && (FUNCTION_WORDS.has(first.lower) || first.lower.length < 3)) continue
    if (words.slice(start, j + 1).some((w) => NAMES_OF_TIME.has(w.lower))) continue
    // Next to an acronym, a figure or a label (UNDP Country Office, the 2025 Annual Report, Annex B, Phase 2).
    if (before && before.lower !== before.text && joinedBySpace(text, before, first) && !isCapitalised(before.text)) continue
    if (/\d\s*$/.test(text.slice(Math.max(0, first.start - 3), first.start))) continue
    if (/^\s+(?:\d|\p{Lu}(?!\p{Ll})|[IVX]+\b)/u.test(text.slice(last.end, last.end + 8))) continue
    out.push({ ip, start: first.start, end: last.end })
  }
  return out
}

function capitalFindings(paragraphs: IndexedParagraph[]): Finding[] {
  const running = paragraphs.filter((ip) => ip.kind === 'body' || ip.kind === 'list')
  const caps = new Map<string, Occurrence[]>()
  for (const ip of running) {
    for (const o of capitalRuns(ip, wordsOf(ip))) {
      const key = keyOf(runText(ip, o))
      caps.set(key, [...(caps.get(key) ?? []), o])
    }
  }
  if (!caps.size) return []
  // The same words in lower case, anywhere outside quotation marks.
  const byFirst = new Map<string, string[][]>()
  for (const key of caps.keys()) {
    const parts = key.split(' ')
    byFirst.set(parts[0] ?? '', [...(byFirst.get(parts[0] ?? '') ?? []), parts])
  }
  const lower = new Map<string, Occurrence[]>()
  for (const ip of running) {
    const words = wordsOf(ip)
    words.forEach((w, i) => {
      const candidates = isLowerCase(w.text) ? byFirst.get(w.lower) : undefined
      if (!candidates || inQuote(ip, w.start)) return
      for (const parts of candidates) {
        const run = words.slice(i, i + parts.length)
        const same = run.length === parts.length && run.every((x, k) => x.lower === parts[k] && isLowerCase(x.text) &&
          (k === 0 || joinedBySpace(ip.p.text, run[k - 1] ?? x, x)))
        const last = run[run.length - 1]
        if (!same || !last) continue
        const key = parts.join(' ')
        lower.set(key, [...(lower.get(key) ?? []), { ip, start: w.start, end: last.end }])
      }
    })
  }
  const out: Finding[] = []
  for (const [key, capital] of caps) {
    const plain = lower.get(key)
    const firstCap = capital[0]
    const firstPlain = plain?.[0]
    if (!plain || !firstCap || !firstPlain) continue
    const odd = capital.length <= plain.length ? firstCap : firstPlain
    out.push({
      id: `consistency:capitals:${key}`, rule: 'consistency/capitals', group: 'consistency', severity: 'copyedit', advisory: true,
      title: `Capitals used two ways: ${q(runText(firstCap.ip, firstCap))} (${capital.length}) and ${q(runText(firstPlain.ip, firstPlain))} (${plain.length})`,
      detail: 'Mid-sentence, capitals are for names. Check whether this is the name of something, and write it one way throughout.',
      excerpt: excerpt(odd.ip.p.text, odd.start, odd.end), location: spanAt(odd.ip.p, odd.start, odd.end),
    })
  }
  return capped(out, 40)
}

// ------------------------------------------------------------------------------------------- compound words
const COMPOUND_FUNCTION_WORDS = new Set(['a', 'an', 'the', 'in', 'on', 'to', 'at', 'by', 'of', 'may', 'any', 'every', 'some', 'all', 'no'])
const PARTICLES = new Set(['up', 'down', 'out', 'off', 'in', 'on', 'over', 'back', 'through', 'away', 'about', 'around'])

/** The singular of a word, roughly: policies to policy, makers to maker; business and analysis stay. */
function singular(w: string): string {
  if (w.length > 4 && w.endsWith('ies')) return `${w.slice(0, -3)}y`
  if (w.length > 4 && w.endsWith('s') && !/(?:ss|us|is)$/.test(w)) return w.slice(0, -1)
  return w
}

type Kind = 'closed' | 'hyphenated' | 'open'
interface Form { kind: Kind; written: string; exact: string; count: number; first: Occurrence }

/**
 * Compounds written closed, hyphenated or open in forms that differ in number as well (policymakers and policy maker).
 * The hyphenation check already pairs forms that match exactly (dataset and data set), so those are left to it.
 */
function compoundFindings(paragraphs: IndexedParagraph[]): Finding[] {
  const prose = paragraphs.filter((ip) => ip.kind !== 'skip' && ip.kind !== 'heading')
  const HYPHEN = /[-\u2010\u2011]/
  // Each candidate in its three kinds, with the normal form (singular, joined up) that groups them.
  const hyphenated: Array<[string, Form]> = []
  const open: Array<[string, Form]> = []
  const closed: Array<[string, Form]> = []
  const form = (kind: Kind, written: string, exact: string, first: Occurrence): Form => ({ kind, written, exact, count: 1, first })
  for (const ip of prose) {
    const words = wordsOf(ip)
    words.forEach((w, i) => {
      if (HYPHEN.test(w.text)) {
        const parts = w.lower.split(HYPHEN)
        const [p1 = '', p2 = ''] = parts
        if (parts.length === 2 && p1.length >= 2 && p2.length >= 2 && !COMPOUND_FUNCTION_WORDS.has(p1) && !PARTICLES.has(p2)) {
          hyphenated.push([singular(p1 + p2), form('hyphenated', w.lower, p1 + p2, { ip, start: w.start, end: w.end })])
        }
        return
      }
      if (w.lower.length >= 5 && !w.lower.includes("'") && !w.lower.includes('\u2019')) {
        closed.push([singular(w.lower), form('closed', w.lower, w.lower, { ip, start: w.start, end: w.end })])
      }
      const b = words[i + 1]
      if (!b || HYPHEN.test(b.text) || !joinedBySpace(ip.p.text, w, b)) return
      if (COMPOUND_FUNCTION_WORDS.has(w.lower) || PARTICLES.has(b.lower) || w.lower.length < 2 || b.lower.length < 2) return
      open.push([singular(w.lower + b.lower), form('open', `${w.lower} ${b.lower}`, w.lower + b.lower, { ip, start: w.start, end: b.end })])
    })
  }
  // Keep only normal forms found in at least two kinds.
  const norms = (list: Array<[string, Form]>): Set<string> => new Set(list.map(([n]) => n))
  const inHyphenated = norms(hyphenated)
  const inOpen = norms(open)
  const inClosed = norms(closed)
  const groups = new Map<string, Map<string, Form>>()
  const keep = (list: Array<[string, Form]>, a: Set<string>, b: Set<string>): void => {
    for (const [norm, f] of list) {
      if (!a.has(norm) && !b.has(norm)) continue
      const forms = groups.get(norm) ?? new Map<string, Form>()
      const known = forms.get(f.written)
      if (known) known.count += 1
      else forms.set(f.written, f)
      groups.set(norm, forms)
    }
  }
  keep(hyphenated, inOpen, inClosed)
  keep(open, inHyphenated, inClosed)
  keep(closed, inHyphenated, inOpen)
  const out: Finding[] = []
  for (const forms of groups.values()) {
    const list = [...forms.values()]
    const kinds = new Set(list.map((f) => f.kind))
    if (kinds.size < 2) continue
    // An exact match between two kinds is the hyphenation check's to report.
    const exactKinds = new Map<string, Set<Kind>>()
    for (const f of list) exactKinds.set(f.exact, (exactKinds.get(f.exact) ?? new Set<Kind>()).add(f.kind))
    if ([...exactKinds.values()].some((k) => k.size > 1)) continue
    list.sort((a, b) => b.count - a.count)
    const [top, odd] = list
    if (!top || !odd) continue
    out.push({
      id: `consistency:compound:${top.written}`, rule: 'consistency/compounds', group: 'consistency', severity: 'copyedit', advisory: true,
      title: `Written more than one way: ${list.map((f) => `${q(f.written)} (${f.count})`).join(', ')}`,
      detail: `Pick one form, closed, hyphenated or open, and use it in the singular and plural alike. The most common is ${q(top.written)}.`,
      excerpt: excerpt(odd.first.ip.p.text, odd.first.start, odd.first.end), location: spanAt(odd.first.ip.p, odd.first.start, odd.first.end),
    })
  }
  return capped(out, 30)
}

// ------------------------------------------------------------------------------------------- accents
/** Words whose plain spelling is a different English word. */
const ACCENT_EXCEPTIONS = new Set(['resume', 'expose', 'rose', 'pate', 'lame', 'sake', 'mate', 'ole'])
const stripAccents = (w: string): string => w.normalize('NFD').replace(/\p{M}/gu, '').normalize('NFC')

/** The same word written with and without accents (cafe with and without its acute). The usual form is offered as a fix. */
function accentFindings(paragraphs: IndexedParagraph[]): Finding[] {
  const prose = paragraphs.filter((ip) => ip.kind !== 'skip')
  const accented = new Set<string>()
  for (const ip of prose) {
    for (const w of wordsOf(ip)) {
      if (/[^\u0000-\u007f]/.test(w.text)) {
        const base = stripAccents(w.text)
        if (base !== w.text && !ACCENT_EXCEPTIONS.has(base.toLowerCase())) accented.add(base)
      }
    }
  }
  if (!accented.size) return []
  const groups = new Map<string, Map<string, Occurrence[]>>()
  for (const ip of prose) {
    for (const w of wordsOf(ip)) {
      const base = stripAccents(w.text)
      if (!accented.has(base) || inQuote(ip, w.start)) continue
      const forms = groups.get(base) ?? new Map<string, Occurrence[]>()
      forms.set(w.text, [...(forms.get(w.text) ?? []), { ip, start: w.start, end: w.end }])
      groups.set(base, forms)
    }
  }
  const out: Finding[] = []
  for (const forms of groups.values()) {
    if (forms.size < 2) continue
    const ranked = [...forms].sort((a, b) => b[1].length - a[1].length)
    const [top, ...rest] = ranked
    if (!top) continue
    const title = `Written two ways: ${ranked.map(([form, list]) => `${q(form)} (${list.length})`).join(' and ')}`
    for (const [form, list] of rest) {
      for (const o of capped(list, 10)) {
        const replacement = matchCase(form, top[0])
        out.push({
          id: `consistency:accent:${o.ip.p.index}:${o.start}`, rule: 'consistency/accents', group: 'consistency', severity: 'copyedit', advisory: true,
          title, detail: `Use one spelling throughout. The most common here is ${q(top[0])}.`,
          excerpt: excerpt(o.ip.p.text, o.start, o.end), location: spanAt(o.ip.p, o.start, o.end),
          ...(stripAccents(replacement) === stripAccents(form) ? { fix: replaceFix(o.ip.p, o.start, o.end, replacement), fixLabel: `Change to ${q(replacement)}` } : {}),
        })
      }
    }
  }
  return out
}

export function termFindings(snapshot: DocSnapshot): Finding[] {
  const paragraphs = indexText(snapshot)
  return [...capitalFindings(paragraphs), ...compoundFindings(paragraphs), ...accentFindings(paragraphs)]
}
