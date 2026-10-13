// Ranges beyond one paragraph's text: paragraph marks, spans of paragraphs, bookmarks, intersections, Word's Find
// codes, and ranges as WordprocessingML for getOoxml and inline insertOoxml.

import { Ctx, FakeObject } from './base.ts'
import type { FakeFont, FakeModel, FakeParagraph } from './model.ts'
import type { Container } from './objects.ts'
import { RangeCollection, RangeObj, trackRanges } from './objects.ts'

/** Word's Find codes the fake understands: ^^ ^l ^t ^s anywhere, and ^p at the end of the text. */
export function parseFind(needle: string): { text: string; mark: boolean } {
  const mark = /(?:^|[^^])\^p$/.test(needle)
  const body = mark ? needle.slice(0, -2) : needle
  const text = body.replace(/\^([\^lts])/g, (_m, c: string) => (c === '^' ? '^' : c === 'l' ? '\u000b' : c === 't' ? '\t' : '\u00a0'))
  return { text, mark }
}

/**
 * Search hits inside text[from, to) of one paragraph; with `whole`, a needle ending in ^p can match the mark. Word's
 * Find never overlaps hits; `overlapping` keeps the older fake behaviour that Paragraph.search has always had here.
 */
export function findInParagraph(
  context: Ctx, container: Container, p: FakeParagraph, needle: string, matchCase: boolean,
  from = 0, to = p.text.length, whole = false, overlapping = false,
): Array<RangeObj | MarkRangeObj> {
  const { text, mark } = parseFind(needle)
  const fold = (s: string): string => (matchCase ? s : s.toLowerCase())
  if (mark) {
    if (!whole || !fold(p.text).endsWith(fold(text))) return []
    return [new MarkRangeObj(context, container, p, text.length)]
  }
  const hay = fold(p.text.slice(0, to))
  const n = fold(text)
  const out: RangeObj[] = []
  for (let i = hay.indexOf(n, from); n && i !== -1; i = hay.indexOf(n, i + (overlapping ? 1 : n.length))) {
    out.push(new RangeObj(context, container, p, i, i + n.length))
  }
  trackRanges(p, out)
  return out
}

/** The end of a paragraph, with `before` characters of its text: what ^p and -^p find. Replacing it joins paragraphs. */
export class MarkRangeObj extends FakeObject {
  constructor(context: Ctx, private readonly container: Container, readonly para: FakeParagraph, private readonly before: number) {
    super(context)
  }
  get text(): string { this.need('text'); return `${this.para.text.slice(this.para.text.length - this.before)}\r` }
  insertText(text: string, location: string): void {
    this.context.queue(() => {
      if (location !== 'Replace') throw new Error(`fake paragraph marks support Replace only (${location})`)
      const list = this.container.list
      const i = list.indexOf(this.para)
      const next = list[i + 1]
      this.para.text = this.para.text.slice(0, this.para.text.length - this.before) + text + (next?.text ?? '')
      if (next) list.splice(i + 1, 1)
    })
  }
}

/** A range over whole paragraphs, from expandTo. Paragraphs merged into an earlier one drop out of it. */
export class SpanRangeObj extends FakeObject {
  constructor(context: Ctx, private readonly container: Container, private readonly paras: FakeParagraph[]) {
    super(context)
  }
  private present(): FakeParagraph[] {
    return this.paras.filter((p) => this.container.list.includes(p))
  }
  get isNullObject(): boolean { return false }
  get text(): string { this.need('text'); return this.present().map((p) => p.text).join('\r') }
  search(needle: string, options: { matchCase?: boolean } = {}): RangeCollection {
    return new RangeCollection(this.context, () =>
      this.present().flatMap((p) => findInParagraph(this.context, this.container, p, needle, Boolean(options.matchCase), 0, p.text.length, true)),
    )
  }
  /** Deleting a span leaves one empty paragraph, as deleting from the start of one paragraph to the end of another does. */
  delete(): void {
    this.context.queue(() => {
      const [first, ...rest] = this.present()
      if (!first) return
      first.text = ''
      const kept = this.container.list.filter((p) => !rest.includes(p))
      this.container.list.splice(0, this.container.list.length, ...kept)
    })
  }
  select(_mode?: string): void {
    this.context.queue(() => {
      const first = this.present()[0]
      if (first) this.container.model.selection = [first]
      this.container.model.selected.push(`span:${first?.text ?? ''}`)
    })
  }
}

export class NullRangeObj extends FakeObject {
  get isNullObject(): boolean { return true }
}

/** The span of a paragraph that a bookmark covers, read when the batch has run. */
export class BookmarkRangeObj extends FakeObject {
  constructor(context: Ctx, private readonly model: FakeModel, private readonly name: string) {
    super(context)
  }
  get isNullObject(): boolean { return !this.model.bookmarks?.has(this.name) }
  get text(): string {
    this.need('text')
    const b = this.model.bookmarks?.get(this.name)
    return b ? b.paragraph.text.slice(b.start, b.end) : ''
  }
}

/** Where `a` lies relative to `b`, for ranges on paragraphs of the same story, as compareLocationWith says it. */
export function compareLocation(a: RangeObj, b: RangeObj): string {
  const list = a.container.list
  const ia = a.para ? list.indexOf(a.para) : -1
  const ib = b.para ? list.indexOf(b.para) : -1
  if (ia < 0 || ib < 0) return 'Unrelated'
  if (ia === ib) return a.start === b.start && a.end === b.end ? 'Equal' : a.end <= b.start ? 'Before' : a.start >= b.end ? 'After' : 'OverlapsBefore'
  if (ia < ib) return ia + 1 === ib ? 'AdjacentBefore' : 'Before'
  return ia === ib + 1 ? 'AdjacentAfter' : 'After'
}

/** The part of `a` that the selection or another range on the same paragraph also covers. */
export function intersect(a: RangeObj, b: RangeObj): RangeObj | NullRangeObj {
  const p = a.para
  if (!p) return new NullRangeObj(a.context)
  const model = b.selectionOf
  let from = b.start
  let to = b.end
  if (model) {
    const span = model.selectionRange
    if (!span) return model.selection.includes(p) ? new RangeObj(a.context, a.container, p, a.start, a.end, null, 'content') : new NullRangeObj(a.context)
    if (span.paragraph !== p) return new NullRangeObj(a.context)
    from = span.start
    to = span.end
  } else if (b.para !== p) return new NullRangeObj(a.context)
  const start = Math.max(a.start, from)
  const end = Math.min(a.end, to)
  return start <= end ? new RangeObj(a.context, a.container, p, start, end, null, 'content') : new NullRangeObj(a.context)
}

/** getTextRanges: the pieces of a range between ending marks, trimmed of spaces when asked. */
export function textRanges(range: RangeObj, marks: string[], trim: boolean): RangeObj[] {
  const p = range.para
  if (!p) return []
  const out: RangeObj[] = []
  let s = range.start
  const flush = (e: number): void => {
    let a = s
    let b = e
    if (trim) {
      while (a < b && /\s/.test(p.text.charAt(a))) a += 1
      while (b > a && /\s/.test(p.text.charAt(b - 1))) b -= 1
    }
    if (b > a) out.push(new RangeObj(range.context, range.container, p, a, b))
  }
  for (let i = range.start; i < range.end; i += 1) {
    if (marks.includes(p.text.charAt(i))) {
      flush(trim ? i : i + 1)
      s = i + 1
    }
  }
  flush(range.end)
  trackRanges(p, out)
  return out
}

const esc = (s: string): string => s.replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;')

export function runProperties(f: FakeFont): string {
  return (
    (f.name ? `<w:rFonts w:ascii="${f.name}" w:hAnsi="${f.name}"/>` : '') +
    (f.bold ? '<w:b/>' : '') +
    (f.italic ? '<w:i/>' : '') +
    (f.color ? `<w:color w:val="${f.color.replace('#', '')}"/>` : '') +
    (f.size ? `<w:sz w:val="${f.size * 2}"/>` : '') +
    (f.highlightColor ? `<w:highlight w:val="${f.highlightColor}"/>` : '')
  )
}

/** A flat OPC package around body XML, closed by the body-level section properties Word always writes. */
export function packageOf(body: string): string {
  return (
    '<pkg:package xmlns:pkg="http://schemas.microsoft.com/office/2006/xmlPackage">' +
    '<pkg:part pkg:name="/word/document.xml"><pkg:xmlData>' +
    `<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:body>${body}<w:sectPr/></w:body></w:document>` +
    '</pkg:xmlData></pkg:part></pkg:package>'
  )
}

/** Text from start to end of a paragraph as one run carrying the paragraph's font, plus its pictures when whole. */
export function paragraphXml(model: FakeModel, p: FakeParagraph, start = 0, end = p.text.length): string {
  const whole = start === 0 && end === p.text.length
  if (whole && p.ooxml) return p.ooxml
  const pictures = whole ? model.pictures.filter((x) => x.paragraph === p).map(() => '<w:r><w:drawing/></w:r>').join('') : ''
  const text = p.text.slice(start, end)
  const run = text ? `<w:r><w:rPr>${runProperties(p.font)}</w:rPr><w:t xml:space="preserve">${esc(text)}</w:t></w:r>` : ''
  return `<w:p><w:pPr><w:pStyle w:val="${esc(p.style)}"/></w:pPr>${run}${pictures}</w:p>`
}

const unesc = (s: string): string => s.replace(/&lt;/g, '<').replace(/&gt;/g, '>').replace(/&quot;/g, '"').replace(/&amp;/g, '&')

/** Reads the text and the first run's formatting of the first paragraph in a package, as Word would apply them. */
export function readInlineOoxml(ooxml: string): { text: string; font: FakeFont } {
  const body = /<w:body>([\s\S]*?)<\/w:body>/.exec(ooxml)?.[1] ?? ooxml
  const p = /<w:p\b[\s\S]*?<\/w:p>/.exec(body)?.[0] ?? ''
  const text = [...p.matchAll(/<w:t(?:\s[^>]*)?>([^<]*)<\/w:t>/g)].map((m) => unesc(m[1] ?? '')).join('')
  const rPr = /<w:r(?:\s[^>]*)?>\s*<w:rPr>([\s\S]*?)<\/w:rPr>/.exec(p)?.[1] ?? ''
  const val = (tag: string): string | null => new RegExp(`<w:${tag}\\b[^>]*?w:(?:val|ascii)="([^"]*)"`).exec(rPr)?.[1] ?? null
  const flag = (tag: string): boolean => new RegExp(`<w:${tag}(?:\\s+w:val="(?!0|false)[^"]*")?\\s*/>`).test(rPr)
  const size = val('sz')
  return {
    text,
    font: {
      name: val('rFonts'),
      size: size ? Number(size) / 2 : null,
      color: val('color'),
      highlightColor: val('highlight'),
      bold: flag('b'),
      italic: flag('i'),
    },
  }
}
