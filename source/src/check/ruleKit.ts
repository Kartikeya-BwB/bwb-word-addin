// Small helpers the consistency rules share: where a finding points, the one-span fix, and quoting in titles.

import type { FixAction, Location, ParagraphInfo } from './types.ts'

export function spanAt(p: ParagraphInfo, start: number, end: number): Location {
  return { kind: 'paragraph', paragraph: p.index, expectedText: p.text, start, end }
}

/** Replaces text[start, end) of one paragraph. The fix checks the paragraph still reads as it did. */
export function replaceFix(p: ParagraphInfo, start: number, end: number, replacement: string): FixAction {
  return { kind: 'replace', paragraph: p.index, expectedText: p.text, start, end, match: p.text.slice(start, end), replacement }
}

/** Typographic quotation marks around a quoted form in a title or detail. */
export const q = (s: string): string => `\u201c${s}\u201d`

/** At most `max` items, so one pattern repeated through a long document cannot flood the list. */
export const capped = <T>(items: T[], max: number): T[] => items.slice(0, max)

/** The form most used; on a tie, the one listed first in `prefer`, then the first seen. */
export function majority<K extends string>(counts: ReadonlyMap<K, number>, prefer: readonly K[] = []): K | null {
  let best: K | null = null
  let bestCount = -1
  for (const [k, n] of counts) {
    const better = n > bestCount || (n === bestCount && best !== null && prefer.indexOf(k) >= 0 && (prefer.indexOf(best) < 0 || prefer.indexOf(k) < prefer.indexOf(best)))
    if (better) {
      best = k
      bestCount = n
    }
  }
  return best
}

/** Gives a replacement the case pattern of what it replaces: UPPER, Capitalised or lower. */
export function matchCase(source: string, replacement: string): string {
  if (source.length > 1 && source === source.toUpperCase() && source !== source.toLowerCase()) return replacement.toUpperCase()
  const first = source.charAt(0)
  if (first !== first.toLowerCase()) return replacement.charAt(0).toUpperCase() + replacement.slice(1)
  return replacement
}
