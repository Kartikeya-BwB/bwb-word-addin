// Copyedit checks a proofreader makes before QA: repeated words, unmatched brackets and
// quotation marks, spaces before punctuation, doubled punctuation and lists that end their items inconsistently.

import type { DocSnapshot, Finding, Location, ParagraphInfo } from './types.ts'
import { excerpt, headingLevel, isVerbatim } from './util.ts'

/** Repeats that are often deliberate ("had had", "that that") are offered as advice, not as errors. */
const OFTEN_DELIBERATE = new Set(['had', 'that', 'is', 'do'])

function at(p: ParagraphInfo, start: number, end: number): Location {
  return { kind: 'paragraph', paragraph: p.index, expectedText: p.text, start, end }
}

function repeatedWords(p: ParagraphInfo): Finding[] {
  const out: Finding[] = []
  for (const m of p.text.matchAll(/\b([\p{L}']+)(\s+)\1\b/giu)) {
    const word = m[1] ?? ''
    const start = m.index
    if (/^\d+$/.test(word) || word.length < 2) continue
    const deliberate = OFTEN_DELIBERATE.has(word.toLowerCase())
    // Remove the second word and the space before it.
    const dupStart = start + word.length
    const dupEnd = start + m[0].length
    out.push({
      id: `copyedit:repeat:${p.index}:${start}`, rule: 'copyedit/repeated-word', group: 'writing', severity: 'copyedit',
      ...(deliberate ? { advisory: true } : {}),
      title: `“${word}” twice in a row`, excerpt: excerpt(p.text, start, dupEnd), location: at(p, start, dupEnd),
      fix: { kind: 'replace', paragraph: p.index, expectedText: p.text, start: dupStart, end: dupEnd, match: p.text.slice(dupStart, dupEnd), replacement: '' },
      fixLabel: 'Remove the repeat',
    })
  }
  return out
}

const PAIRS: Array<[string, string, string]> = [['(', ')', 'bracket'], ['[', ']', 'square bracket'], ['“', '”', 'quotation mark']]

function unmatched(p: ParagraphInfo): Finding[] {
  const out: Finding[] = []
  for (const [open, close, name] of PAIRS) {
    let depth = 0
    let firstBad = -1
    for (let i = 0; i < p.text.length; i++) {
      const ch = p.text[i]
      if (ch === open) depth += 1
      else if (ch === close) {
        depth -= 1
        if (depth < 0 && firstBad < 0) firstBad = i
      }
    }
    // A list item such as "a) first point" closes a bracket it never opened on purpose.
    const listMarker = open === '(' && /^\s*[a-z0-9ivx]{1,4}\)\s/i.test(p.text)
    if ((depth !== 0 || firstBad >= 0) && !listMarker) {
      const where = firstBad >= 0 ? firstBad : p.text.lastIndexOf(open)
      out.push({
        id: `copyedit:unmatched:${p.index}:${name}`, rule: 'copyedit/unmatched', group: 'writing', severity: 'copyedit',
        title: `Unmatched ${name}`, detail: `Every ${open} needs a ${close} in the same paragraph.`,
        excerpt: excerpt(p.text, Math.max(0, where), Math.max(0, where) + 1), location: at(p, Math.max(0, where), Math.max(0, where) + 1),
      })
    }
  }
  return out
}

function punctuation(p: ParagraphInfo): Finding[] {
  const out: Finding[] = []
  for (const m of p.text.matchAll(/(?<=[\p{L}\p{N}”)\]])( +)([,;:.?])(?=\s|$)/gu)) {
    const start = m.index
    const end = start + (m[1]?.length ?? 0)
    out.push({
      id: `copyedit:space-before:${p.index}:${start}`, rule: 'copyedit/space-before-punctuation', group: 'writing', severity: 'copyedit',
      title: `Space before “${m[2] ?? ''}”`, excerpt: excerpt(p.text, start, end + 1), location: at(p, start, end + 1),
      fix: { kind: 'replace', paragraph: p.index, expectedText: p.text, start, end, match: p.text.slice(start, end), replacement: '' },
      fixLabel: 'Remove the space',
    })
  }
  for (const m of p.text.matchAll(/(?<!\.)([,;:])\1+|(?<![.…])\.\.(?!\.)/gu)) {
    const start = m.index
    const end = start + m[0].length
    out.push({
      id: `copyedit:double:${p.index}:${start}`, rule: 'copyedit/doubled-punctuation', group: 'writing', severity: 'copyedit',
      title: `Doubled punctuation “${m[0]}”`, excerpt: excerpt(p.text, start, end), location: at(p, start, end),
      fix: { kind: 'replace', paragraph: p.index, expectedText: p.text, start, end, match: m[0], replacement: m[0][0] ?? '' },
      fixLabel: `Change to “${m[0][0] ?? ''}”`,
    })
  }
  return out
}

/** Runs of list items in the same style whose items end some with a full stop and some without. */
function listEndings(snapshot: DocSnapshot): Finding[] {
  const out: Finding[] = []
  const ps = snapshot.paragraphs
  let i = 0
  while (i < ps.length) {
    const first = ps[i]
    if (!first?.isListItem || first.inTable) {
      i += 1
      continue
    }
    const run: ParagraphInfo[] = []
    while (i < ps.length && ps[i]?.isListItem && ps[i]?.style === first.style && !ps[i]?.inTable) {
      const p = ps[i]
      if (p) run.push(p)
      i += 1
    }
    const items = run.filter((p) => p.text.trim())
    if (items.length < 3) continue
    const withStop = items.filter((p) => /[.!?]["”)]?$/.test(p.text.trim()))
    if (withStop.length > 0 && withStop.length < items.length) {
      const odd = withStop.length * 2 < items.length ? withStop : items.filter((p) => !withStop.includes(p))
      const p = odd[0]
      if (!p) continue
      out.push({
        id: `copyedit:list-endings:${items[0]?.index ?? 0}`, rule: 'copyedit/list-endings', group: 'consistency', severity: 'copyedit', advisory: true,
        title: `List items end inconsistently (${withStop.length} of ${items.length} with a full stop)`,
        detail: 'End every item of a list the same way: full stops for sentences, none for fragments.',
        excerpt: excerpt(p.text, 0, 60), location: { kind: 'paragraph', paragraph: p.index, expectedText: p.text },
      })
    }
  }
  return out
}

export function copyeditFindings(snapshot: DocSnapshot): Finding[] {
  const out: Finding[] = []
  for (const p of snapshot.paragraphs) {
    if (!p.text.trim() || isVerbatim(p)) continue
    out.push(...repeatedWords(p), ...punctuation(p))
    if (headingLevel(p) === null) out.push(...unmatched(p))
  }
  return [...out, ...listEndings(snapshot)]
}
