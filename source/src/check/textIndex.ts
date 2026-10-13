// What the consistency checks need to know about each paragraph, worked out once per snapshot: what kind of text it
// is, where its sentences start and which parts are quoted. Quoted words, pull quotes, contents entries and writing
// prompts are someone else's words or generated text, so the checks leave them alone.

import type { DocSnapshot, ParagraphInfo } from './types.ts'
import { headingLevel, isTitle, isVerbatim } from './util.ts'

/** body: running text; list: list items; table: cells; heading: headings and titles; caption: captions; skip: not checked. */
export type ParagraphKind = 'body' | 'list' | 'table' | 'heading' | 'caption' | 'skip'

export interface IndexedParagraph {
  p: ParagraphInfo
  kind: ParagraphKind
  /** Offsets where a sentence (or a line, cell or clause after a colon) starts. */
  starts: ReadonlySet<number>
  /** Spans [start, end) inside quotation marks. */
  quotes: ReadonlyArray<readonly [number, number]>
}

const CAPTION_LABEL = /^\s*(?:Figure|Fig\.?|Table|Chart|Graph|Map|Exhibit|Box)\s+\d+/i
/** Abbreviations after which a full stop does not end a sentence (compared in lower case). */
const ABBREVIATIONS = new Set([
  'e.g', 'i.e', 'etc', 'cf', 'vs', 'approx', 'no', 'nos', 'fig', 'figs', 'p', 'pp', 'para', 'paras', 'vol', 'ed', 'eds',
  'mr', 'mrs', 'ms', 'dr', 'prof', 'st', 'ltd', 'inc', 'co', 'jan', 'feb', 'mar', 'apr', 'jun', 'jul', 'aug', 'sep',
  'sept', 'oct', 'nov', 'dec', 'al', 'est', 'incl', 'excl', 'min', 'max',
])
const OPENERS = /["'(\[\u2018\u201c\u2022\u25aa\u25a0\u2013\u2014*\s-]/

export function kindOf(p: ParagraphInfo, guidance: ReadonlySet<string>): ParagraphKind {
  if (!p.text.trim() || isVerbatim(p) || guidance.has(p.text.trim()) || /^table of figures$/i.test(p.style)) return 'skip'
  if (headingLevel(p) !== null || isTitle(p) || /title|subtitle|cover|divider|heading/i.test(p.style)) return 'heading'
  if (/caption/i.test(p.style) || p.styleBuiltIn === 'Caption' || CAPTION_LABEL.test(p.text)) return 'caption'
  if (p.inTable) return 'table'
  if (p.isListItem) return 'list'
  return 'body'
}

/** Where sentences start: the first word of the text, and the first word after . ! ? : or a line break. */
export function sentenceStarts(text: string): Set<number> {
  const out = new Set<number>()
  const firstAfter = (i: number): number => {
    let j = i
    while (j < text.length && OPENERS.test(text.charAt(j))) j += 1
    return j
  }
  out.add(firstAfter(0))
  for (const m of text.matchAll(/[.!?:]["'\u2019\u201d)\]]*(?=\s|$)|[\u000b\t\n]/g)) {
    const at = m.index
    if (text.charAt(at) === '.') {
      const word = /([\p{L}.]+)$/u.exec(text.slice(Math.max(0, at - 12), at))?.[1]?.toLowerCase() ?? ''
      if (ABBREVIATIONS.has(word) || /^\p{L}$/u.test(word)) continue
    }
    out.add(firstAfter(at + m[0].length))
  }
  return out
}

/** Spans inside quotation marks: typographic double and single quotes, and straight double quotes in pairs. */
export function quotedSpans(text: string): Array<[number, number]> {
  const out: Array<[number, number]> = []
  let open = -1
  let closer = ''
  for (let i = 0; i < text.length; i += 1) {
    const ch = text.charAt(i)
    if (open < 0) {
      const before = text.charAt(i - 1)
      if (ch === '\u201c') [open, closer] = [i, '\u201d']
      else if (ch === '"') [open, closer] = [i, '"']
      else if (ch === '\u2018' && !/[\p{L}\p{N}]/u.test(before)) [open, closer] = [i, '\u2019']
    } else if (ch === closer && (closer !== '\u2019' || !/[\p{L}\p{N}]/u.test(text.charAt(i + 1)))) {
      out.push([open, i + 1])
      open = -1
    }
  }
  if (open >= 0) out.push([open, text.length])
  return out
}

const cache = new WeakMap<DocSnapshot, IndexedParagraph[]>()

/** The paragraphs of a snapshot with their kind, sentence starts and quoted spans. Worked out once per snapshot. */
export function indexText(snapshot: DocSnapshot): IndexedParagraph[] {
  const known = cache.get(snapshot)
  if (known) return known
  const guidance = new Set(snapshot.controls.filter((c) => c.tag === 'BwBGuidance').map((c) => c.text.trim()))
  const out = snapshot.paragraphs.map((p): IndexedParagraph => {
    const kind = kindOf(p, guidance)
    return kind === 'skip'
      ? { p, kind, starts: new Set(), quotes: [] }
      : { p, kind, starts: sentenceStarts(p.text), quotes: quotedSpans(p.text) }
  })
  cache.set(snapshot, out)
  return out
}

export const inQuote = (ip: IndexedParagraph, at: number): boolean => ip.quotes.some(([a, b]) => at >= a && at < b)

/** True when the word at `at` starts a sentence, so its capital says nothing about the word. */
export const atSentenceStart = (ip: IndexedParagraph, at: number): boolean => ip.starts.has(at)
