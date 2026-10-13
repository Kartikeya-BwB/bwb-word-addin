// Fake text frames, text ranges, fonts and paragraph formats. Text is kept as runs, so writing a substring keeps the
// formatting of the text around it, as PowerPoint does.

import { Ctx, FakeObject, FakeOfficeError } from './base.ts'
import type { FakeFont, FakePresentation, FakeRun, FakeShape, FakeTextFrame } from './model.ts'
import { font } from './model.ts'

export function frameText(f: FakeTextFrame): string {
  return f.runs.map((r) => r.text).join('')
}

/** Splits runs so a run starts at `offset`; returns that run's index. */
function splitAt(f: FakeTextFrame, offset: number): number {
  let pos = 0
  for (let i = 0; i < f.runs.length; i += 1) {
    const r = f.runs[i]
    if (!r) break
    if (offset === pos) return i
    if (offset < pos + r.text.length) {
      f.runs.splice(i, 1, { text: r.text.slice(0, offset - pos), font: { ...r.font } }, { text: r.text.slice(offset - pos), font: { ...r.font } })
      return i + 1
    }
    pos += r.text.length
  }
  return f.runs.length
}

/** Merges neighbouring runs with the same font and drops empty runs, keeping at least one. */
function tidy(f: FakeTextFrame): void {
  const out: FakeRun[] = []
  for (const r of f.runs) {
    const last = out[out.length - 1]
    if (last && JSON.stringify(last.font) === JSON.stringify(r.font)) last.text += r.text
    else if (r.text || !out.length) out.push(r)
  }
  const first = out[0]
  if (first && !first.text && out.length > 1) out.shift()
  f.runs = out.length ? out : [{ text: '', font: font() }]
}

/** The font of the character at `offset`, or of the last character before it. */
export function fontAt(f: FakeTextFrame, offset: number): FakeFont {
  let pos = 0
  let last: FakeFont | undefined
  for (const r of f.runs) {
    if (r.text && offset < pos + r.text.length) return r.font
    if (r.text || !last) last = r.font
    pos += r.text.length
  }
  return last ?? font()
}

export function runsBetween(f: FakeTextFrame, start: number, end: number): FakeRun[] {
  const a = splitAt(f, start)
  const b = splitAt(f, end)
  return f.runs.slice(a, b)
}

/** Replaces [start, end) with text in the font of the first replaced character. */
export function replaceText(f: FakeTextFrame, start: number, end: number, text: string): void {
  const style = { ...fontAt(f, start) }
  const a = splitAt(f, start)
  const b = splitAt(f, end)
  f.runs.splice(a, b - a, { text, font: style })
  tidy(f)
}

export function writeFont(f: FakeTextFrame, start: number, end: number, values: Partial<FakeFont>): void {
  if (end <= start) return
  runsBetween(f, start, end).forEach((r) => Object.assign(r.font, values))
  tidy(f)
}

/** The value shared by every run in [start, end), or null when they differ, as Office.js reports mixed text. */
export function readFont<K extends keyof FakeFont>(f: FakeTextFrame, start: number, end: number, key: K): FakeFont[K] | null {
  if (end <= start) return fontAt(f, start)[key]
  const copy: FakeTextFrame = { ...f, runs: f.runs.map((r) => ({ text: r.text, font: { ...r.font } })) }
  const values = runsBetween(copy, start, end).map((r) => r.font[key])
  const first = values[0]
  return first !== undefined && values.every((v) => v === first) ? first : null
}

export function normaliseColour(value: string): string {
  const hex = value.trim().replace(/^#/, '')
  return /^[0-9a-f]{6}$/i.test(hex) ? `#${hex.toUpperCase()}` : value
}

/** What a text range needs from its shape: the frame (failing the sync when the shape has none) and the deck. */
export interface FrameSource {
  model: FakePresentation
  shape: () => FakeShape
  /** The shape's proxy, for TextFrame.getParentShape(). */
  owner?: () => FakeObject
}

function frameOf(source: FrameSource): FakeTextFrame {
  const s = source.shape()
  if (!s.textFrame) throw new FakeOfficeError('InvalidArgument', `Shape ${s.id} (${s.type}) has no text frame`)
  return s.textFrame
}

export class FontObj extends FakeObject {
  constructor(context: Ctx, private readonly span: () => { frame: FakeTextFrame; start: number; end: number }) {
    super(context)
  }
  private get<K extends keyof FakeFont>(key: K): FakeFont[K] | null {
    this.need(key)
    const s = this.span()
    return readFont(s.frame, s.start, s.end, key)
  }
  private set<K extends keyof FakeFont>(key: K, value: FakeFont[K]): void {
    this.write(() => {
      const s = this.span()
      writeFont(s.frame, s.start, s.end, { [key]: key === 'color' && typeof value === 'string' ? normaliseColour(value) : value })
    })
  }
  get name(): string | null { return this.get('name') }
  set name(v: string | null) { this.set('name', v) }
  get size(): number | null { return this.get('size') }
  set size(v: number | null) { this.set('size', v) }
  get color(): string | null { return this.get('color') }
  set color(v: string | null) { this.set('color', v) }
  get bold(): boolean | null { return this.get('bold') }
  set bold(v: boolean | null) { this.set('bold', v) }
  get italic(): boolean | null { return this.get('italic') }
  set italic(v: boolean | null) { this.set('italic', v) }
  get underline(): string | null { return this.get('underline') }
  set underline(v: string | null) { this.set('underline', v) }
  /** PowerPointApi 1.8: true, false, or null for text that is partly superscript. */
  get superscript(): boolean | null {
    this.need('superscript')
    const s = this.span()
    const copy: FakeTextFrame = { ...s.frame, runs: s.frame.runs.map((r) => ({ text: r.text, font: { ...r.font } })) }
    const values = s.end > s.start ? runsBetween(copy, s.start, s.end).map((r) => r.font.superscript ?? false) : [fontAt(s.frame, s.start).superscript ?? false]
    return values.every((v) => v === values[0]) ? values[0] ?? false : null
  }
  set superscript(v: boolean | null) {
    this.context.gate('1.8', 'ShapeFont.superscript')
    this.write(() => {
      const s = this.span()
      writeFont(s.frame, s.start, s.end, { superscript: v === true })
    })
  }
  protected override versionFor(path: string): string | undefined {
    return path === 'superscript' ? '1.8' : undefined
  }
}

/** Bullets for the whole frame: the fake keeps one visible flag per text frame (bulletVisible). */
export class BulletFormatObj extends FakeObject {
  constructor(context: Ctx, private readonly frame: () => FakeTextFrame) {
    super(context)
  }
  get visible(): boolean { this.need('visible'); return this.frame().bulletVisible }
  set visible(v: boolean) { this.write(() => (this.frame().bulletVisible = v)) }
}

export class ParagraphFormatObj extends FakeObject {
  private bullets: BulletFormatObj | undefined
  constructor(context: Ctx, private readonly frame: () => FakeTextFrame) {
    super(context)
  }
  get horizontalAlignment(): string { this.need('horizontalAlignment'); return this.frame().horizontalAlignment }
  set horizontalAlignment(v: string) { this.write(() => (this.frame().horizontalAlignment = v)) }
  get bulletFormat(): BulletFormatObj { return (this.bullets ??= new BulletFormatObj(this.context, this.frame)) }
  protected override onLoaded(head: string, rest: string): void {
    if (head === 'bulletFormat') this.bulletFormat.markLoaded(rest || '*')
  }
}

export class TextRangeObj extends FakeObject {
  private fontObj: FontObj | undefined
  private paragraphObj: ParagraphFormatObj | undefined

  /** `length` null runs to the end of the text. */
  constructor(context: Ctx, private readonly source: FrameSource, private readonly from: number, private len: number | null) {
    super(context)
  }

  private span(): { frame: FakeTextFrame; start: number; end: number } {
    const frame = frameOf(this.source)
    const total = frameText(frame).length
    const end = this.len === null ? total : this.from + this.len
    if (this.from < 0 || end < this.from || end > total) {
      throw new FakeOfficeError('InvalidArgument', `Text range ${this.from} to ${end} is outside the text (length ${total})`)
    }
    return { frame, start: this.from, end }
  }

  protected override checkLoad(): void {
    this.span()
  }

  protected override versionFor(path: string): string | undefined {
    return path === 'start' || path === 'length' ? '1.5' : undefined
  }

  get text(): string {
    this.need('text')
    const s = this.span()
    return frameText(s.frame).slice(s.start, s.end)
  }
  set text(v: string) {
    this.write(() => {
      const s = this.span()
      replaceText(s.frame, s.start, s.end, v)
      this.len = v.length
    })
  }
  get start(): number { this.need('start'); return this.span().start }
  get length(): number { this.need('length'); const s = this.span(); return s.end - s.start }

  get font(): FontObj {
    this.fontObj ??= new FontObj(this.context, () => this.span())
    return this.fontObj
  }
  get paragraphFormat(): ParagraphFormatObj {
    this.paragraphObj ??= new ParagraphFormatObj(this.context, () => this.span().frame)
    return this.paragraphObj
  }

  protected override onLoaded(head: string, rest: string): void {
    if (head === 'font') this.font.markLoaded(rest || '*')
    if (head === 'paragraphFormat') this.paragraphFormat.markLoaded(rest || '*')
  }

  getSubstring(start: number, length?: number): TextRangeObj {
    const len = length ?? (this.len === null ? null : this.len - start)
    return new TextRangeObj(this.context, this.source, this.from + start, len)
  }

  setSelected(): void {
    this.context.gate('1.5', 'TextRange.setSelected')
    this.write(() => {
      const s = this.span()
      const shape = this.source.shape()
      this.source.model.selection = { ...this.source.model.selection, shapeIds: [shape.id], textRange: { shapeId: shape.id, start: s.start, length: s.end - s.start } }
    })
  }

  getParentTextFrame(): TextFrameObj {
    this.context.gate('1.5', 'TextRange.getParentTextFrame')
    return new TextFrameObj(this.context, this.source)
  }
}

export class TextFrameObj extends FakeObject {
  private range: TextRangeObj | undefined

  constructor(context: Ctx, private readonly source: FrameSource, private readonly nullable = false) {
    super(context)
  }

  protected override isNull(): boolean {
    return this.nullable && !this.source.shape().textFrame
  }

  protected override checkLoad(): void {
    if (!this.nullable) frameOf(this.source)
  }

  private frame(): FakeTextFrame {
    return frameOf(this.source)
  }
  private prop<K extends keyof FakeTextFrame>(key: K): FakeTextFrame[K] {
    this.need(key)
    return this.frame()[key]
  }
  private put<K extends keyof FakeTextFrame>(key: K, value: FakeTextFrame[K]): void {
    this.write(() => {
      this.frame()[key] = value
    })
  }

  get autoSizeSetting(): string { return this.prop('autoSizeSetting') }
  set autoSizeSetting(v: string) { this.put('autoSizeSetting', v) }
  get verticalAlignment(): string { return this.prop('verticalAlignment') }
  set verticalAlignment(v: string) { this.put('verticalAlignment', v) }
  get wordWrap(): boolean { return this.prop('wordWrap') }
  set wordWrap(v: boolean) { this.put('wordWrap', v) }
  get leftMargin(): number { return this.prop('leftMargin') }
  set leftMargin(v: number) { this.put('leftMargin', v) }
  get rightMargin(): number { return this.prop('rightMargin') }
  set rightMargin(v: number) { this.put('rightMargin', v) }
  get topMargin(): number { return this.prop('topMargin') }
  set topMargin(v: number) { this.put('topMargin', v) }
  get bottomMargin(): number { return this.prop('bottomMargin') }
  set bottomMargin(v: number) { this.put('bottomMargin', v) }
  get hasText(): boolean { this.need('hasText'); return frameText(this.frame()).length > 0 }

  get textRange(): TextRangeObj {
    this.range ??= new TextRangeObj(this.context, this.source, 0, null)
    return this.range
  }

  protected override onLoaded(head: string, rest: string): void {
    if (head === 'textRange') this.textRange.markLoaded(rest || '*')
  }

  deleteText(): void {
    this.write(() => {
      const f = this.frame()
      f.runs = [{ text: '', font: { ...fontAt(f, 0) } }]
    })
  }

  /** PowerPointApi 1.5: the shape that holds this text frame. */
  getParentShape(): FakeObject {
    this.context.gate('1.5', 'TextFrame.getParentShape')
    const owner = this.source.owner?.()
    if (!owner) throw new FakeOfficeError('GeneralException', 'This text frame has no parent shape in the fake')
    return owner
  }
}
