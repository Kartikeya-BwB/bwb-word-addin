import { readFileSync } from 'node:fs'
import { describe, expect, it } from 'vitest'
import {
  codePointAt, codePointBefore, cpLength, firstCodePoint, moveByCodePoints, PY_IGNORECASE_ASCII_LETTER, PY_SPACE,
  PY_WORD, pyDecimalValue, pyIgnoreCase, pyInt, pyIsAlnum, pyIsDigit, pyIsUpper, pyRstrip,
} from './writingUnicode.ts'

type Ranges = Array<[number, number]>
interface Tables {
  unidata: string
  assigned: Ranges
  w: Ranges
  d: Ranges
  s: Ranges
  isdigit: Ranges
  isalnum: Ranges
  isupper: Ranges
  decimalZeros: number[]
  ignoreCase: Record<string, number[]>
  ignoreCaseAsciiClass: number[]
  strings: Array<{ s: string; isupper: boolean; upper: string; lower: string }>
}
const fixture = JSON.parse(
  readFileSync(new URL('../../test/fixtures/writing-expected.json', import.meta.url), 'utf8'),
) as { unicode: Tables }
const py = fixture.unicode
const MAX = 0x10ffff

function bitmap(ranges: Ranges): Uint8Array {
  const out = new Uint8Array(MAX + 1)
  for (const [a, b] of ranges) out.fill(1, a, b + 1)
  return out
}
const assigned = bitmap(py.assigned)

/** Code points where the JavaScript predicate and Python's table disagree. */
function divergence(ranges: Ranges, predicate: (ch: string) => boolean, last = MAX): number[] {
  const want = bitmap(ranges)
  const out: number[] = []
  for (let cp = 0; cp <= last; cp++) {
    if ((want[cp] === 1) !== predicate(String.fromCodePoint(cp))) out.push(cp)
  }
  return out
}
const hex = (cps: number[]): string[] => cps.map((cp) => `U+${cp.toString(16).toUpperCase().padStart(4, '0')}`)

describe(`Python ${py.unidata} character classes and str predicates`, () => {
  const word = new RegExp(`^[${PY_WORD}]$`, 'u')
  const digit = /^\p{Nd}$/u
  const space = new RegExp(`^[${PY_SPACE}]$`, 'u')
  const checks: Array<[string, Ranges, (ch: string) => boolean]> = [
    ['\\w', py.w, (ch) => word.test(ch)],
    ['\\d', py.d, (ch) => digit.test(ch)],
    ['str.isdigit()', py.isdigit, pyIsDigit],
    ['str.isalnum()', py.isalnum, pyIsAlnum],
    ['str.isupper()', py.isupper, pyIsUpper],
  ]
  it.each(checks)('%s agrees with Python on every code point Python has assigned', (name, ranges, predicate) => {
    const diff = divergence(ranges, predicate)
    const onAssigned = diff.filter((cp) => assigned[cp] === 1)
    console.info(`${name}: ${diff.length} code points differ, all unassigned in Unicode ${py.unidata}`)
    expect(hex(onAssigned)).toEqual([])
  })

  it('\\s is exactly Python\'s set', () => {
    expect(hex(divergence(py.s, (ch) => space.test(ch)))).toEqual([])
  })

  it('reads decimal digits of every script as int() does', () => {
    const zeros = [...py.decimalZeros].sort((a, b) => a - b)
    const wrong: number[] = []
    for (const [a, b] of py.d) {
      for (let cp = a; cp <= b; cp++) {
        const zero = zeros.filter((z) => z <= cp).pop() ?? 0
        if (pyDecimalValue(String.fromCodePoint(cp)) !== cp - zero) wrong.push(cp)
      }
    }
    expect(hex(wrong)).toEqual([])
    expect(pyInt('1\u0662\u{1D7D1}')).toBe(123)
    expect(pyInt('007')).toBe(7)
  })

  it('matches letters case-insensitively exactly as re.I does', () => {
    for (const [letter, cps] of Object.entries(py.ignoreCase)) {
      const re = new RegExp(`^${pyIgnoreCase(letter)}$`, 'u')
      const got: number[] = []
      for (let cp = 0; cp <= 0xffff; cp++) if (re.test(String.fromCodePoint(cp))) got.push(cp)
      expect(hex(got), letter).toEqual(hex(cps))
    }
    const cls = new RegExp(`^${PY_IGNORECASE_ASCII_LETTER}$`, 'u')
    const got: number[] = []
    for (let cp = 0; cp <= 0xffff; cp++) if (cls.test(String.fromCodePoint(cp))) got.push(cp)
    expect(hex(got)).toEqual(hex(py.ignoreCaseAsciiClass))
  })

  it('upper-cases, lower-cases and tests isupper() on sample strings as Python does', () => {
    for (const x of py.strings) {
      expect([x.s, pyIsUpper(x.s), x.s.toUpperCase(), x.s.toLowerCase()]).toEqual([x.s, x.isupper, x.upper, x.lower])
    }
  })
})

describe('string helpers', () => {
  const smile = '\u{1F600}'
  it('strips Python whitespace only', () => {
    expect(pyRstrip('a \t\u3000\u0085\u001C')).toBe('a')
    expect(pyRstrip('a\uFEFF')).toBe('a\uFEFF')
    expect(pyRstrip('   ')).toBe('')
  })

  it('counts and steps over code points', () => {
    expect(cpLength(`a${smile}b`)).toBe(3)
    expect(firstCodePoint(`${smile}x`)).toBe(smile)
    expect(firstCodePoint('')).toBe('')
    expect(codePointBefore(`a${smile}`, 3)).toBe(smile)
    expect(codePointBefore('ab', 0)).toBe('')
    expect(codePointAt(`a${smile}`, 1)).toBe(smile)
    expect(codePointAt('a', 1)).toBe('')
    expect(moveByCodePoints(`${smile}${smile}x`, 5, -2)).toBe(2)
    expect(moveByCodePoints(`${smile}${smile}x`, 0, 2)).toBe(4)
    expect(moveByCodePoints('abc', 1, 25)).toBe(3)
  })
})
