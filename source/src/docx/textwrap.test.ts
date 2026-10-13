import { describe, expect, it } from 'vitest'
import { wrap } from './textwrap.ts'

// Expected lines from Python 3.13: textwrap.wrap(text, width)
const CASES: Array<[string, number, string[]]> = [
  ['Blue Carbon Finance for Pacific Fisheries', 25, ['Blue Carbon Finance for', 'Pacific Fisheries']],
  ['Look, goof-ball -- use the -b option!', 10, ['Look,', 'goof-ball', '-- use the', '-b option!']],
  [
    'A very-long-hyphenated-compound-word-that-must-break here',
    12,
    ['A very-long-', 'hyphenated-', 'compound-', 'word-that-', 'must-break', 'here'],
  ],
  ['Supercalifragilisticexpialidocious is long', 10, ['Supercalif', 'ragilistic', 'expialidoc', 'ious is', 'long']],
  ['tabs\there\tand\nnewlines\r\nand  double  spaces', 9, ['tabs', 'here', 'and', 'newlines', 'and', 'double', 'spaces']],
  ['   leading and trailing   ', 8, ['leading', 'and', 'trailing']],
  ['\u{DC}n\u{EF}c\u{F6}d\u{E9} \u{1F30D} Title with Wide Characters', 12, ['\u{DC}n\u{EF}c\u{F6}d\u{E9} \u{1F30D}', 'Title with', 'Wide', 'Characters']],
  ['', 10, []],
  ['x---y--z a--b', 4, ['x---', 'y--z', 'a--b']],
  ['well-known co-operation pre-2020 e-mail', 7, ['well-', 'known', 'co-oper', 'ation p', 're-2020', 'e-mail']],
]

describe('wrap', () => {
  it.each(CASES)('wraps %j at %i like Python', (text, width, lines) => {
    expect(wrap(text, width)).toEqual(lines)
  })

  it('refuses a width below 1', () => {
    expect(() => wrap('x', 0)).toThrow(RangeError)
  })
})
