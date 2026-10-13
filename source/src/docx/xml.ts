/**
 * Shared constants and small XML helpers for the BwB document format: a port of bwb_lib/common.py.
 *
 * The builder must produce the same WordprocessingML as the Python builder, byte for byte, so the helpers below also
 * reproduce the Python behaviours the output depends on: round() rounds half to even, len() and slicing count code
 * points, and sum() of floats uses Neumaier compensation.
 */

// ---- Brand palette (BwB Brand Guidelines v2.0) ----
export const NAVY = '022338'
export const GREEN = '328840'
export const LIME = 'BED62F'
export const CHARCOAL = '4A4A4C'
export const GREEN50 = '7CC287'
export const GREEN20 = 'EBF1EB'
export const CHAR50 = '929294'
export const CHAR20 = 'D5D5D6'
export const LIME60 = 'D4E56E'
export const LIME20 = 'EEF5B8'
export const LIGHT = 'E8E9E6'
export const OFFWHITE = 'F7F8F5'
export const BLACK = '000000'
export const WHITE = 'FFFFFF'

// ---- Page geometry (A4, twips unless stated) ----
export const PAGE_W = 11906
export const PAGE_H = 16838
export const MARGIN = 1224
export const TEXT_W = PAGE_W - 2 * MARGIN
export const EMU_PER_TWIP = 635
export const PAGE_W_EMU = 7560310
export const PAGE_H_EMU = 10692130

export const HEAD_FONT = 'Candara'
export const BODY_FONT = 'Aptos'

export const W_NS = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main'
export const R_NS = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships'

export const ROOT_NS =
  'xmlns:wpc="http://schemas.microsoft.com/office/word/2010/wordprocessingCanvas" ' +
  'xmlns:cx="http://schemas.microsoft.com/office/drawing/2014/chartex" ' +
  'xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006" ' +
  'xmlns:aink="http://schemas.microsoft.com/office/drawing/2016/ink" ' +
  'xmlns:am3d="http://schemas.microsoft.com/office/drawing/2017/model3d" ' +
  'xmlns:o="urn:schemas-microsoft-com:office:office" ' +
  'xmlns:oel="http://schemas.microsoft.com/office/2019/extlst" ' +
  'xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" ' +
  'xmlns:m="http://schemas.openxmlformats.org/officeDocument/2006/math" ' +
  'xmlns:v="urn:schemas-microsoft-com:vml" ' +
  'xmlns:wp14="http://schemas.microsoft.com/office/word/2010/wordprocessingDrawing" ' +
  'xmlns:wp="http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing" ' +
  'xmlns:w10="urn:schemas-microsoft-com:office:word" ' +
  'xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" ' +
  'xmlns:w14="http://schemas.microsoft.com/office/word/2010/wordml" ' +
  'xmlns:w15="http://schemas.microsoft.com/office/word/2012/wordml" ' +
  'xmlns:w16cex="http://schemas.microsoft.com/office/word/2018/wordml/cex" ' +
  'xmlns:w16cid="http://schemas.microsoft.com/office/word/2016/wordml/cid" ' +
  'xmlns:w16="http://schemas.microsoft.com/office/word/2018/wordml" ' +
  'xmlns:w16sdtdh="http://schemas.microsoft.com/office/word/2020/wordml/sdtdatahash" ' +
  'xmlns:w16se="http://schemas.microsoft.com/office/word/2015/wordml/symex" ' +
  'xmlns:wpg="http://schemas.microsoft.com/office/word/2010/wordprocessingGroup" ' +
  'xmlns:wpi="http://schemas.microsoft.com/office/word/2010/wordprocessingInk" ' +
  'xmlns:wne="http://schemas.microsoft.com/office/word/2006/wordml" ' +
  'xmlns:wps="http://schemas.microsoft.com/office/word/2010/wordprocessingShape" ' +
  'mc:Ignorable="w14 w15 w16se w16cid w16 w16cex w16sdtdh wp14"'

export const XML_DECL = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n'

// Characters XML 1.0 cannot carry (control characters, unpaired surrogates). Python writes them through and Word then
// refuses the file, so they are dropped here; text without them is escaped exactly as Python does.
const INVALID_XML =
  /[\u{0}-\u{8}\u{B}\u{C}\u{E}-\u{1F}\u{D800}-\u{DFFF}\u{FFFE}\u{FFFF}]/gu

/** xml.sax.saxutils.escape plus double quotes, as common.esc. */
export function esc(s: string | number): string {
  return String(s)
    .replace(INVALID_XML, '')
    .replace(/&/g, '&amp;')
    .replace(/>/g, '&gt;')
    .replace(/</g, '&lt;')
    .replace(/"/g, '&quot;')
}

// ---- Python number and string semantics ----

/** Python's round(x) for a float: nearest integer, halves to even. */
export function pyRound(x: number): number {
  const f = Math.floor(x)
  if (x - f === 0.5) return f % 2 === 0 ? f : f + 1
  return Math.round(x)
}

/** Python's round(x, nd) for a float: the exact binary value rounded to nd decimals, exact halves to even. */
export function pyRoundTo(x: number, nd: number): number {
  const exact = Math.abs(x).toFixed(100)
  const dot = exact.indexOf('.')
  const rest = exact.slice(dot + 1 + nd)
  if (/^50*$/.test(rest)) {
    const digits = (exact.slice(0, dot) + exact.slice(dot + 1, dot + 1 + nd)).replace(/^0+(?=\d)/, '')
    let n = BigInt(digits)
    if (n % 2n === 1n) n += 1n
    const s = n.toString().padStart(nd + 1, '0')
    const v = Number(nd ? `${s.slice(0, s.length - nd)}.${s.slice(s.length - nd)}` : s)
    return x < 0 ? -v : v
  }
  return Number(x.toFixed(nd))
}

// Characters for which Python's str.isspace() is true; str.strip(), str.split() and \s in re patterns use this set.
export const PY_SPACE = '\\t\\n\\x0b\\x0c\\r\\x1c-\\x1f \\x85\\xa0\\u1680\\u2000-\\u200a\\u2028\\u2029\\u202f\\u205f\\u3000'
const PY_STRIP = new RegExp(`^[${PY_SPACE}]+|[${PY_SPACE}]+$`, 'g')
const PY_RSTRIP = new RegExp(`[${PY_SPACE}]+$`)
const PY_SPLIT = new RegExp(`[${PY_SPACE}]+`)

/** Python's re.split(r"\s+", s): empty strings at either end are kept. */
export function pyReSplitWs(s: string): string[] {
  return s.split(PY_SPLIT)
}

/** Python's str.strip() with no argument. */
export function pyStrip(s: string): string {
  return s.replace(PY_STRIP, '')
}

/** Python's str.rstrip() with no argument. */
export function pyRstrip(s: string): string {
  return s.replace(PY_RSTRIP, '')
}

/** Python's str.split() with no argument: runs of whitespace separate words, empty words are dropped. */
export function pySplit(s: string): string[] {
  return s.split(PY_SPLIT).filter((w) => w !== '')
}

/** Python's str.strip(chars) for a set of single characters. */
export function pyStripChars(s: string, chars: string): string {
  let a = 0
  let b = s.length
  while (a < b && chars.includes(s.charAt(a))) a++
  while (b > a && chars.includes(s.charAt(b - 1))) b--
  return s.slice(a, b)
}

/** Number of code points (Python's len of a str). */
export function cpLen(s: string): number {
  let n = 0
  for (let i = 0; i < s.length; i++) {
    const c = s.charCodeAt(i)
    if (c >= 0xd800 && c <= 0xdbff && i + 1 < s.length && (s.charCodeAt(i + 1) & 0xfc00) === 0xdc00) i++
    n++
  }
  return n
}

/** The first n code points of s (Python's s[:n]). */
export function cpHead(s: string, n: number): string {
  if (s.length <= n) return s
  let i = 0
  let k = 0
  while (i < s.length && k < n) {
    const c = s.charCodeAt(i)
    i += c >= 0xd800 && c <= 0xdbff && i + 1 < s.length && (s.charCodeAt(i + 1) & 0xfc00) === 0xdc00 ? 2 : 1
    k++
  }
  return s.slice(0, i)
}

/** str.replace(find, repl, 1) without JavaScript's $ patterns in the replacement. */
export function replaceFirst(s: string, find: string, repl: string): string {
  const i = s.indexOf(find)
  return i < 0 ? s : s.slice(0, i) + repl + s.slice(i + find.length)
}

/** A number that remembers whether Python would hold it as an int or a float (it changes how sum() adds). */
export interface PyNumber {
  readonly v: number
  readonly isInt: boolean
}

/**
 * Python 3.12+ sum(): ints are added exactly until the first float; from then on floats are added with Neumaier
 * compensation (applied once at the end) while ints are added without it.
 */
export function pySum(items: readonly PyNumber[]): PyNumber {
  let i = 0
  let iResult = 0
  while (i < items.length && (items[i] as PyNumber).isInt) {
    iResult += (items[i] as PyNumber).v
    i++
  }
  if (i === items.length) return { v: iResult, isInt: true }
  let f = iResult + (items[i] as PyNumber).v
  let c = 0
  for (i += 1; i < items.length; i++) {
    const item = items[i] as PyNumber
    if (item.isInt) {
      f += item.v
      continue
    }
    const x = item.v
    const t = f + x
    if (Math.abs(f) >= Math.abs(x)) c += f - t + x
    else c += x - t + f
    f = t
  }
  if (c !== 0 && Number.isFinite(c)) f += c
  return { v: f, isInt: false }
}

/** Python's max(a, b, ...): the first of the largest values, keeping its int or float kind. */
export function pyMax(...items: PyNumber[]): PyNumber {
  let best = items[0] as PyNumber
  for (const it of items.slice(1)) if (it.v > best.v) best = it
  return best
}

// ---- runs and paragraphs ----

export interface RunProps {
  b?: boolean | null
  i?: boolean | null
  color?: string | null
  sz?: number | null
  font?: string | null
  rstyle?: string | null
  vert?: string | null
  caps?: boolean
  u?: string | null
  noproof?: boolean
  spacing?: number | null
}

/** Build a <w:rPr> element. sz is in points. */
export function rpr(p: RunProps = {}): string {
  const parts: string[] = []
  if (p.rstyle) parts.push(`<w:rStyle w:val="${p.rstyle}"/>`)
  if (p.font) parts.push(`<w:rFonts w:ascii="${p.font}" w:eastAsia="${p.font}" w:hAnsi="${p.font}" w:cs="${p.font}"/>`)
  if (p.b === true) parts.push('<w:b/><w:bCs/>')
  else if (p.b === false) parts.push('<w:b w:val="0"/><w:bCs w:val="0"/>')
  if (p.i === true) parts.push('<w:i/><w:iCs/>')
  else if (p.i === false) parts.push('<w:i w:val="0"/><w:iCs w:val="0"/>')
  if (p.caps) parts.push('<w:caps/>')
  if (p.noproof) parts.push('<w:noProof/>')
  if (p.color) parts.push(`<w:color w:val="${p.color}"/>`)
  if (p.spacing !== undefined && p.spacing !== null) parts.push(`<w:spacing w:val="${p.spacing}"/>`)
  if (p.sz) {
    const hp = pyRound(p.sz * 2)
    parts.push(`<w:sz w:val="${hp}"/><w:szCs w:val="${hp}"/>`)
  }
  if (p.u) parts.push(`<w:u w:val="${p.u}"/>`)
  if (p.vert) parts.push(`<w:vertAlign w:val="${p.vert}"/>`)
  return parts.length ? `<w:rPr>${parts.join('')}</w:rPr>` : ''
}

/** A text run; tabs and line breaks inside the text become their own runs. */
export function run(text: string, p: RunProps = {}): string {
  const rp = rpr(p)
  let out = ''
  for (const piece of String(text).split(/(\t|\n)/)) {
    if (piece === '\t') out += `<w:r>${rp}<w:tab/></w:r>`
    else if (piece === '\n') out += `<w:r>${rp}<w:br/></w:r>`
    else if (piece !== '') out += `<w:r>${rp}<w:t xml:space="preserve">${esc(piece)}</w:t></w:r>`
  }
  return out
}

/** Complex field with a cached result. */
export function field(instr: string, cached: string, p: RunProps = {}): string {
  const rp = rpr(p)
  return (
    `<w:r>${rp}<w:fldChar w:fldCharType="begin"/></w:r>` +
    `<w:r>${rp}<w:instrText xml:space="preserve"> ${esc(instr)} </w:instrText></w:r>` +
    `<w:r>${rp}<w:fldChar w:fldCharType="separate"/></w:r>` +
    `<w:r>${rp}<w:t>${esc(cached)}</w:t></w:r>` +
    `<w:r>${rp}<w:fldChar w:fldCharType="end"/></w:r>`
  )
}

/** Attribute map for w:spacing and w:ind; insertion order is the attribute order. */
export type AttrMap = Readonly<Record<string, string | number>>

export interface ParaProps {
  style?: string | null
  jc?: string | null
  keep_next?: boolean
  keep_lines?: boolean
  page_break_before?: boolean
  num?: readonly [number, number] | null
  spacing?: AttrMap | null
  ind?: AttrMap | null
  tabs?: readonly string[] | null
  extra?: string
  sect?: string | null
}

function attrs(m: AttrMap): string {
  return Object.entries(m)
    .map(([k, v]) => `w:${k}="${v}"`)
    .join(' ')
}

/** Build <w:pPr>; the child order follows the CT_PPr schema sequence. */
export function ppr(p: ParaProps = {}): string {
  const parts: string[] = []
  let extra = p.extra ?? ''
  if (p.style) parts.push(`<w:pStyle w:val="${p.style}"/>`)
  if (p.keep_next) parts.push('<w:keepNext/>')
  if (p.keep_lines) parts.push('<w:keepLines/>')
  if (p.page_break_before) parts.push('<w:pageBreakBefore/>')
  if (p.num) parts.push(`<w:numPr><w:ilvl w:val="${p.num[1]}"/><w:numId w:val="${p.num[0]}"/></w:numPr>`)
  if (extra && extra.startsWith('<w:pBdr')) {
    parts.push(extra)
    extra = ''
  }
  if (p.tabs && p.tabs.length) parts.push('<w:tabs>' + p.tabs.join('') + '</w:tabs>')
  if (p.spacing && Object.keys(p.spacing).length) parts.push(`<w:spacing ${attrs(p.spacing)}/>`)
  if (p.ind && Object.keys(p.ind).length) parts.push(`<w:ind ${attrs(p.ind)}/>`)
  if (p.jc) parts.push(`<w:jc w:val="${p.jc}"/>`)
  if (extra) parts.push(extra)
  if (p.sect) parts.push(p.sect)
  return parts.length ? `<w:pPr>${parts.join('')}</w:pPr>` : ''
}

export function para(content = '', p: ParaProps = {}): string {
  return `<w:p>${ppr(p)}${content}</w:p>`
}

/** Empty paragraph with an exact height (twips). */
export function spacer(twips: number, after = 0): string {
  return (
    `<w:p><w:pPr><w:spacing w:before="0" w:after="${after}" w:line="${twips}" w:lineRule="exact"/>` +
    '<w:rPr><w:sz w:val="2"/><w:szCs w:val="2"/></w:rPr></w:pPr></w:p>'
  )
}

// ---- Inline mini mark-up: **bold**, *italic*, [label](url), ^[footnote text] ----

const TOKEN = /(\^\[.+?\]|\*\*.+?\*\*|(?<!\*)\*(?!\*).+?(?<!\*)\*(?!\*)|\[[^\]]+\]\([^)]+\))/gs
const LINK = /^\[([^\]]+)\]\(([^)]+)\)/
const ESC_MAP: ReadonlyArray<readonly [string, string]> = [
  ['\\*', '\u0001'],
  ['\\[', '\u0002'],
  ['\\^', '\u0003'],
  ['\\]', '\u0004'],
  ['\\\\', '\u0005'],
]
const UNESC: Readonly<Record<string, string>> = { '\u0001': '*', '\u0002': '[', '\u0003': '^', '\u0004': ']', '\u0005': '\\' }

export type InlineKind = 'text' | 'bold' | 'italic' | 'link' | 'footnote'

export interface InlineToken {
  readonly kind: InlineKind
  readonly value: string
  /** The URL or #bookmark of a link; null for the other kinds. */
  readonly extra: string | null
}

/** Escape characters that the mini mark-up would otherwise treat as formatting. */
export function escapeInline(text: string): string {
  return String(text)
    .replaceAll('\\', '\\\\')
    .replaceAll('*', '\\*')
    .replaceAll('[', '\\[')
    .replaceAll(']', '\\]')
    .replaceAll('^', '\\^')
}

function unesc(s: string): string {
  return s.replace(/[\u0001-\u0005]/g, (c) => UNESC[c] ?? c)
}

/** Split text into text, bold, italic, link and footnote tokens. A backslash escapes * [ ] ^ and itself. */
export function splitInline(text: string): InlineToken[] {
  let t = text
  for (const [k, v] of ESC_MAP) t = t.split(k).join(v)
  const out: InlineToken[] = []
  let pos = 0
  for (const m of t.matchAll(TOKEN)) {
    const start = m.index
    if (start > pos) out.push({ kind: 'text', value: unesc(t.slice(pos, start)), extra: null })
    const tok = m[0]
    if (tok.startsWith('^[')) out.push({ kind: 'footnote', value: unesc(tok.slice(2, -1)), extra: null })
    else if (tok.startsWith('**')) out.push({ kind: 'bold', value: unesc(tok.slice(2, -2)), extra: null })
    else if (tok.startsWith('*')) out.push({ kind: 'italic', value: unesc(tok.slice(1, -1)), extra: null })
    else {
      const mm = LINK.exec(tok)
      out.push({ kind: 'link', value: unesc(mm?.[1] ?? ''), extra: unesc(mm?.[2] ?? '') })
    }
    pos = start + tok.length
  }
  if (pos < t.length) out.push({ kind: 'text', value: unesc(t.slice(pos)), extra: null })
  return out
}
