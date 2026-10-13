// Pure rules for the text tools: rejoining lines pasted from a PDF, non-breaking spaces that keep figures together,
// and the symbols palette.

export type JoinAction = 'keep' | 'space' | 'join' | 'dehyphenate'

/** A break between two lines: a paragraph mark, or the nth manual line break inside a paragraph. */
export interface LineBreak {
  kind: 'paragraph' | 'line'
  paragraph: number
  /** For a manual line break: which one in the paragraph (0 first), and which among those after a hyphen. */
  index: number
  hyphenIndex: number
  action: JoinAction
}

const SENTENCE_END = /[.!?:;]["'\u2019\u201d)\]]*\s*$/
const LIST_START = /^\s*(?:[\u2022\u25cf\u25a0\u25aa\u25e6\u2023\u2013\u2014*-]\s|\(?\d{1,3}[.)]\s|\(?[a-z][.)]\s|\(?[ivx]{1,4}[.)]\s)/i
const HYPHENATED = /\p{L}-$/u

/** What to do with the break between two lines, given the longest line in the selection. */
export function joinAction(before: string, after: string, longest: number): JoinAction {
  if (!before.trim() || !after.trim()) return 'keep'
  if (SENTENCE_END.test(before) || LIST_START.test(after)) return 'keep'
  // A line well short of the others ends a paragraph or is a heading, not a sentence broken by the page width.
  if (longest >= 40 && before.trim().length < 0.6 * longest) return 'keep'
  if (HYPHENATED.test(before)) return /^\p{Ll}/u.test(after) ? 'dehyphenate' : 'join'
  return /\s$/.test(before) || /^\s/.test(after) ? 'join' : 'space'
}

/** Every break between the lines of the given paragraphs (whose text has manual line breaks as \v) and what to do. */
export function planJoins(paragraphs: string[]): LineBreak[] {
  const lines = paragraphs.map((t) => t.split('\u000b'))
  const longest = Math.max(0, ...lines.flat().map((l) => l.trim().length))
  const out: LineBreak[] = []
  lines.forEach((parts, p) => {
    let hyphens = 0
    parts.forEach((line, k) => {
      const next = k + 1 < parts.length ? parts[k + 1] : lines[p + 1]?.[0]
      if (next === undefined) return
      const kind = k + 1 < parts.length ? 'line' : 'paragraph'
      const hyphenIndex = kind === 'line' && line.endsWith('-') ? hyphens++ : -1
      out.push({ kind, paragraph: p, index: kind === 'line' ? k : -1, hyphenIndex, action: joinAction(line, next, longest) })
    })
  })
  return out
}

/** The text the plan gives, with runs of spaces collapsed: what joining does to the paragraphs, as plain strings. */
export function applyJoins(paragraphs: string[]): string[] {
  const plan = planJoins(paragraphs)
  const out: string[] = []
  let current = ''
  paragraphs.forEach((text, p) => {
    const parts = text.split('\u000b')
    parts.forEach((line, k) => {
      current += line
      const b = plan.find((x) => x.paragraph === p && (k + 1 < parts.length ? x.index === k : x.kind === 'paragraph'))
      if (!b || b.action === 'keep') {
        if (k + 1 < parts.length) current += '\u000b'
        else {
          out.push(current)
          current = ''
        }
      } else if (b.action === 'space') current += ' '
      else if (b.action === 'dehyphenate') current = current.slice(0, -1)
    })
  })
  if (current) out.push(current)
  return out.map(collapseSpaces)
}

export const collapseSpaces = (text: string): string => text.replace(/ {2,}/g, ' ')

export interface Span {
  start: number
  end: number
  replacement: string
}

const MONTHS = 'January|February|March|April|May|June|July|August|September|October|November|December|Jan|Feb|Mar|Apr|Jun|Jul|Aug|Sept?|Oct|Nov|Dec'
const UNITS = [
  'km\u00b2', 'km2', 'km', 'cm', 'mm', 'm\u00b2', 'm2', 'm\u00b3', 'm3', 'm', 'ha', 'kg', 'g', 'kt', 'Mt', 'Gt', 't',
  'tCO2e', 'tCO\u2082e', 'MtCO2e', 'MtCO\u2082e', 'GtCO2e', 'GtCO\u2082e', 'tCO2', 'kWh', 'MWh', 'GWh', 'TWh', 'Wh',
  'kWp', 'MWp', 'GWp', 'kW', 'MW', 'GW', 'TW', 'W', 'kV', 'MJ', 'GJ', 'PJ', 'L', 'ml', 'bbl', '\u00b0C',
  'million', 'billion', 'trillion', 'thousand',
].sort((a, b) => b.length - a.length).map((u) => u.replace(/[.*+?^${}()|[\]\\]/g, '\\$&')).join('|')

/** Spaces that should not break, as positions in the text. */
function keepTogether(text: string): number[] {
  const at = new Set<number>()
  const add = (re: RegExp, groups: number[]): void => {
    for (const m of text.matchAll(re)) {
      // Each listed group is a single space; find where it sits from the groups before it.
      let pos = m.index
      for (let g = 1; g < m.length; g += 1) {
        const part = m[g] ?? ''
        if (groups.includes(g) && part === ' ') at.add(pos)
        pos += part.length
      }
    }
  }
  // USD 5m, EUR 1.2bn: a currency code then an amount.
  add(/(?<![\p{L}\p{N}])([A-Z]{3})( )([-\u2212(]?\d)/gu, [2])
  // 12 per cent, with a space or none already between per and cent.
  add(/(\d)( )(per)([ \u00a0])(cent)(?![\p{L}])/gu, [2, 4])
  // 5 October 2026.
  add(new RegExp(`(?<![\\p{L}\\p{N}])(\\d{1,2})( )(${MONTHS})(\\.?)( )(\\d{4})(?!\\p{N})`, 'gu'), [2, 5])
  // 5 km, 20 MW, 2.5 million.
  add(new RegExp(`(\\d)( )(${UNITS})(?![\\p{L}\\p{N}])`, 'gu'), [2])
  return [...at].sort((a, b) => a - b)
}

/**
 * The replacements that put non-breaking spaces into USD 5m, 12 per cent, 5 October 2026 and 20 MW. Each span covers
 * the words either side of its spaces, so a search for it finds that place and no other.
 */
export function nonBreakingSpans(text: string): Span[] {
  const spans: Span[] = []
  for (const p of keepTogether(text)) {
    let start = p
    while (start > 0 && text.charAt(start - 1) !== ' ') start -= 1
    let end = p + 1
    while (end < text.length && text.charAt(end) !== ' ') end += 1
    const last = spans[spans.length - 1]
    if (last && start <= last.end) last.end = Math.max(last.end, end)
    else spans.push({ start, end, replacement: '' })
  }
  const nbsp = new Set(keepTogether(text))
  for (const s of spans) {
    s.replacement = [...text.slice(s.start, s.end)].map((ch, i) => (nbsp.has(s.start + i) ? '\u00a0' : ch)).join('')
  }
  return spans
}

/** The symbols palette. The en dash is for ranges in tables only; running text writes ranges with "to". */
export const SYMBOLS: ReadonlyArray<{ char: string; label: string }> = [
  { char: '\u2713', label: 'Tick' },
  { char: '\u2717', label: 'Cross' },
  { char: '\u2192', label: 'Arrow right' },
  { char: '\u2190', label: 'Arrow left' },
  { char: '\u2191', label: 'Arrow up' },
  { char: '\u2193', label: 'Arrow down' },
  { char: '\u2022', label: 'Bullet' },
  { char: '\u00b7', label: 'Middle dot' },
  { char: '\u00b0', label: 'Degree' },
  { char: '\u20ac', label: 'Euro' },
  { char: '\u00a3', label: 'Pound' },
  { char: '\u00b1', label: 'Plus or minus' },
  { char: '\u00d7', label: 'Multiplied by' },
  { char: '\u2264', label: 'Less than or equal to' },
  { char: '\u2265', label: 'Greater than or equal to' },
  { char: '\u2248', label: 'Approximately' },
  { char: '\u2026', label: 'Ellipsis' },
  { char: '\u2018', label: 'Opening single quote' },
  { char: '\u2019', label: 'Closing single quote or apostrophe' },
  { char: '\u201c', label: 'Opening double quote' },
  { char: '\u201d', label: 'Closing double quote' },
]

export const EN_DASH = '\u2013'
