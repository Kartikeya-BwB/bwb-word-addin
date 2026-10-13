import { describe, expect, it } from 'vitest'
import { LUCIDE_NAMES } from '../../test/helpers/iconRenderer.ts'
import { normaliseIconName } from './iconPolicy.ts'
import { LUCIDE_ALIASES } from './lucideAliases.ts'
import { getCloseMatches } from './similar.ts'

describe('getCloseMatches', () => {
  it('matches the difflib documentation examples', () => {
    expect(getCloseMatches('appel', ['ape', 'apple', 'peach', 'puppy'])).toEqual(['apple', 'ape'])
    // equal scores are ordered by the larger word first, as heapq.nlargest orders (score, word) tuples
    expect(getCloseMatches('wheel', ['while', 'wheal', 'whee', 'wheels'], 4)).toEqual(['wheels', 'whee', 'wheal', 'while'])
  })

  it('gives the same icon matches as Python over the Lucide names', () => {
    const pool = [...LUCIDE_NAMES, ...Object.keys(LUCIDE_ALIASES)]
    // expected values from difflib.get_close_matches(normalise(word), pool, n=3, cutoff=0.6) in Python 3.13
    const expected: Record<string, string[]> = {
      'no-such-icon': [],
      'not-an-icon': ['octagon'],
      leef: ['leaf', 'beef', 'delete'],
      chart: ['heart', 'cat', 'car'],
      alert: ['wallet', 'tablet', 'trailer'],
      'dolar-sign': ['dollar-sign', 'badge-dollar-sign', 'circle-dollar-sign'],
      x: ['x'],
      user: ['user', 'users', 'user-x'],
      'arrow-right-circle': ['arrow-right-circle', 'arrow-right-left', 'arrow-up-right-from-circle'],
      barchart: ['bar-chart', 'bar-chart-4', 'bar-chart-3'],
    }
    for (const [word, want] of Object.entries(expected)) expect(getCloseMatches(normaliseIconName(word), pool, 3, 0.6), word).toEqual(want)
  })

  it('rejects n below 1', () => {
    expect(() => getCloseMatches('a', ['a'], 0)).toThrow(RangeError)
  })
})
