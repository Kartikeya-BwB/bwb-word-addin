/**
 * Port of Python's textwrap.wrap with its default options (expand tabs, replace whitespace, drop whitespace, break long
 * words, break on hyphens). The builder only needs the number of lines a cover title takes, but that number decides the
 * title size, so the chunking follows TextWrapper.wordsep_re exactly, with Python's Unicode \w and \d.
 */
import { cpLen, pyStrip } from './xml.ts'

const W = '[\\p{L}\\p{N}_]'
const WORD_PUNCT = '[\\p{L}\\p{N}_!"\'&.,?]'
const LETTER = '[\\p{L}\\p{Nl}\\p{No}_]'
const WS = '[\\t\\n\\x0b\\x0c\\r ]'
const NWS = '[^\\t\\n\\x0b\\x0c\\r ]'

const WORDSEP = new RegExp(
  `(${WS}+` +
    `|(?<=${WORD_PUNCT})-{2,}(?=${W})` +
    `|${NWS}+?(?:` +
    `-(?:(?<=${LETTER}{2}-)|(?<=${LETTER}-${LETTER}-))(?=${LETTER}-?${LETTER})` +
    `|(?=${WS}|$)` +
    `|(?<=${WORD_PUNCT})(?=-{2,}${W})` +
    '))',
  'u',
)

function expandTabs(s: string, tabsize = 8): string {
  let out = ''
  let col = 0
  for (const ch of s) {
    if (ch === '\t') {
      const n = tabsize - (col % tabsize)
      out += ' '.repeat(n)
      col += n
    } else if (ch === '\n' || ch === '\r') {
      out += ch
      col = 0
    } else {
      out += ch
      col++
    }
  }
  return out
}

function splitChunks(text: string): string[] {
  const munged = expandTabs(text).replace(/[\t\n\x0b\x0c\r]/g, ' ')
  return munged.split(WORDSEP).filter((c) => c !== '')
}

function isBlank(s: string): boolean {
  return pyStrip(s) === ''
}

function handleLongWord(reversed: string[], curLine: string[], curLen: number, width: number): void {
  const spaceLeft = width < 1 ? 1 : width - curLen
  if (spaceLeft > 0) {
    const chunk = Array.from(reversed[reversed.length - 1] ?? '')
    let end = spaceLeft
    if (chunk.length > spaceLeft) {
      const hyphen = chunk.slice(0, spaceLeft).lastIndexOf('-')
      if (hyphen > 0 && chunk.slice(0, hyphen).some((c) => c !== '-')) end = hyphen + 1
    }
    curLine.push(chunk.slice(0, end).join(''))
    reversed[reversed.length - 1] = chunk.slice(end).join('')
  } else if (!curLine.length) {
    curLine.push(reversed.pop() ?? '')
  }
}

/** textwrap.wrap(text, width): the lines, each at most width code points. */
export function wrap(text: string, width: number): string[] {
  if (width <= 0) throw new RangeError(`invalid width ${width} (must be > 0)`)
  const chunks = splitChunks(text).reverse()
  const lines: string[] = []
  while (chunks.length) {
    const curLine: string[] = []
    let curLen = 0
    if (lines.length && isBlank(chunks[chunks.length - 1] ?? '')) chunks.pop()
    while (chunks.length) {
      const l = cpLen(chunks[chunks.length - 1] ?? '')
      if (curLen + l <= width) {
        curLine.push(chunks.pop() ?? '')
        curLen += l
      } else break
    }
    if (chunks.length && cpLen(chunks[chunks.length - 1] ?? '') > width) handleLongWord(chunks, curLine, curLen, width)
    if (curLine.length && isBlank(curLine[curLine.length - 1] ?? '')) curLine.pop()
    if (curLine.length) lines.push(curLine.join(''))
  }
  return lines
}
