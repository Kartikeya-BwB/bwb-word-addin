/**
 * Python 3 `re` and `str` behaviour that the writing rules rely on, rebuilt for JavaScript.
 *
 * Python's `re` is Unicode-aware on str patterns: `\w` is str.isalnum() plus "_", `\d` is any decimal digit (Nd),
 * `\s` is str.isspace() and `\b` is a boundary between `\w` and non-`\w`. In JavaScript `\w`, `\d` and `\b` are
 * ASCII-only even with the u flag, and `\s` is a different set (it has U+FEFF, lacks U+001C to U+001F and U+0085).
 * So the patterns are built from the class bodies below and compiled with the u flag. With re.I, Python also lets
 * "i" match U+0130 and U+0131, which JavaScript's i flag does not, so case-insensitive parts are spelled out per letter
 * (pyIgnoreCase) and the patterns do not use the i flag at all.
 *
 * Indexes in Python strings count code points; JavaScript counts UTF-16 code units. The helpers that look at "the
 * character before" or "the next character" work on whole code points so that emoji and other astral characters
 * behave as in Python.
 */

/** Body of a character class equal to Python's `\w` (str.isalnum() or "_"). */
export const PY_WORD = '\\p{L}\\p{N}_'
/** Body of a character class equal to Python's `\s` (str.isspace()). */
export const PY_SPACE =
  '\\t\\n\\x0B\\x0C\\r\\x1C-\\x1F \\x85\\xA0\\u1680\\u2000-\\u200A\\u2028\\u2029\\u202F\\u205F\\u3000'
/** Python's `\d`. */
export const PY_DIGIT = '\\p{Nd}'
/** Python's `\b`. */
export const PY_BOUNDARY = `(?:(?<=[${PY_WORD}])(?![${PY_WORD}])|(?<![${PY_WORD}])(?=[${PY_WORD}]))`
/** Python's `[A-Za-z]` under re.I: the ASCII letters plus U+0130, U+0131, U+017F (long s) and U+212A (Kelvin). */
export const PY_IGNORECASE_ASCII_LETTER = '[A-Za-z\\u0130\\u0131\\u017F\\u212A]'

// Characters that Python's re.I treats as equal to an ASCII letter besides its two cases.
const IGNORECASE_EXTRA: Readonly<Record<string, string>> = { i: '\\u0130\\u0131', k: '\\u212A', s: '\\u017F' }
const SYNTAX_CHARS = '^$\\.*+?()[]{}|/'

/** The regex source for a literal matched as Python matches it under re.I. */
export function pyIgnoreCase(literal: string): string {
  let out = ''
  for (const ch of literal) {
    const lower = ch.toLowerCase()
    if (/^[A-Za-z]$/.test(ch)) {
      out += `[${lower}${lower.toUpperCase()}${IGNORECASE_EXTRA[lower] ?? ''}]`
    } else {
      out += SYNTAX_CHARS.includes(ch) ? `\\${ch}` : ch
    }
  }
  return out
}

/** A character class of the given ASCII letters as Python matches it under re.I (`[sd]` with re.I). */
export function pyIgnoreCaseClass(letters: string): string {
  let body = ''
  for (const ch of letters) {
    const lower = ch.toLowerCase()
    body += `${lower}${lower.toUpperCase()}${IGNORECASE_EXTRA[lower] ?? ''}`
  }
  return `[${body}]`
}

// ------------------------------------------------------------------------------------------- str predicates
// str.isdigit() is wider than \d: it also takes superscripts, circled digits and other Numeric_Type=Digit characters.
const DIGIT_NOT_DECIMAL =
  '\\xB2\\xB3\\xB9\\u1369-\\u1371\\u19DA\\u2070\\u2074-\\u2079\\u2080-\\u2089\\u2460-\\u2468\\u2474-\\u247C' +
  '\\u2488-\\u2490\\u24EA\\u24F5-\\u24FD\\u24FF\\u2776-\\u277E\\u2780-\\u2788\\u278A-\\u2792' +
  '\\u{10A40}-\\u{10A43}\\u{10E60}-\\u{10E68}\\u{11052}-\\u{1105A}\\u{1F100}-\\u{1F10A}'
const RE_ISDIGIT = new RegExp(`^[\\p{Nd}${DIGIT_NOT_DECIMAL}]$`, 'u')
const RE_ISALNUM = /^[\p{L}\p{N}]$/u
const RE_DECIMAL = /^\p{Nd}$/u
const RE_UPPER = /\p{Uppercase}/u
const RE_LOWER_OR_TITLE = /[\p{Lowercase}\p{Lt}]/u

/** str.isdigit() for one character (a code point string); false for the empty string. */
export function pyIsDigit(ch: string): boolean {
  return RE_ISDIGIT.test(ch)
}

/** str.isalnum() for one character (a code point string); false for the empty string. */
export function pyIsAlnum(ch: string): boolean {
  return RE_ISALNUM.test(ch)
}

/** str.isupper(): no lower-case or title-case character, and at least one upper-case one. */
export function pyIsUpper(s: string): boolean {
  let cased = false
  for (const ch of s) {
    if (RE_LOWER_OR_TITLE.test(ch)) return false
    if (!cased && RE_UPPER.test(ch)) cased = true
  }
  return cased
}

/** str.isspace() for one UTF-16 code unit (every Python whitespace character is in the BMP). */
function isPySpaceUnit(c: number): boolean {
  return (c >= 0x09 && c <= 0x0d) || (c >= 0x1c && c <= 0x20) || c === 0x85 || c === 0xa0 || c === 0x1680 ||
    (c >= 0x2000 && c <= 0x200a) || c === 0x2028 || c === 0x2029 || c === 0x202f || c === 0x205f || c === 0x3000
}

/** str.rstrip() with no argument. */
export function pyRstrip(s: string): string {
  let end = s.length
  while (end > 0 && isPySpaceUnit(s.charCodeAt(end - 1))) end--
  return s.slice(0, end)
}

/** The value of one decimal digit of any script. Decimal digits are encoded in runs of ten, starting at zero. */
export function pyDecimalValue(ch: string): number {
  const cp = ch.codePointAt(0) ?? 0
  if (cp >= 0x30 && cp <= 0x39) return cp - 0x30
  let start = cp
  while (start > 0 && RE_DECIMAL.test(String.fromCodePoint(start - 1))) start--
  return (cp - start) % 10
}

/** int(s) for a string of decimal digits, as Python reads it (any script, leading zeros allowed). */
export function pyInt(digits: string): number {
  let n = 0
  for (const ch of digits) n = n * 10 + pyDecimalValue(ch)
  return n
}

// ------------------------------------------------------------------------------------------- code points
function isHigh(c: number): boolean {
  return c >= 0xd800 && c <= 0xdbff
}

function isLow(c: number): boolean {
  return c >= 0xdc00 && c <= 0xdfff
}

/** len(s) in Python: the number of code points. */
export function cpLength(s: string): number {
  let n = 0
  for (let i = 0; i < s.length; i++) {
    if (!(isLow(s.charCodeAt(i)) && i > 0 && isHigh(s.charCodeAt(i - 1)))) n++
  }
  return n
}

/** s[:1] in Python: the first code point, or "" for the empty string. */
export function firstCodePoint(s: string): string {
  const cp = s.codePointAt(0)
  return cp === undefined ? '' : String.fromCodePoint(cp)
}

/** The code point that ends at UTF-16 offset i, or "" at the start. */
export function codePointBefore(s: string, i: number): string {
  if (i <= 0) return ''
  if (i >= 2 && isLow(s.charCodeAt(i - 1)) && isHigh(s.charCodeAt(i - 2))) return s.slice(i - 2, i)
  return s.slice(i - 1, i)
}

/** The code point that starts at UTF-16 offset i, or "" at the end. */
export function codePointAt(s: string, i: number): string {
  return i >= s.length ? '' : firstCodePoint(s.slice(i, i + 2))
}

/** Moves a UTF-16 offset by `count` code points (negative moves left), stopping at the ends of the string. */
export function moveByCodePoints(s: string, i: number, count: number): number {
  let pos = i
  if (count < 0) {
    for (let k = 0; k < -count && pos > 0; k++) {
      pos -= pos >= 2 && isLow(s.charCodeAt(pos - 1)) && isHigh(s.charCodeAt(pos - 2)) ? 2 : 1
    }
  } else {
    for (let k = 0; k < count && pos < s.length; k++) {
      pos += pos + 1 < s.length && isHigh(s.charCodeAt(pos)) && isLow(s.charCodeAt(pos + 1)) ? 2 : 1
    }
  }
  return pos
}
