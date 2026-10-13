// Fake Office.js objects for paragraphs and ranges. Tables, content controls, pictures, ranges across paragraphs and
// review objects live in their own files; the classes tests use are re-exported here.

import { Ctx, FakeCollection, FakeObject, clientResult } from './base.ts'
import type { FakeControl, FakeFont, FakeModel, FakeParagraph, FakeTable } from './model.ts'
import { ingestOoxml, para } from './model.ts'
import { ControlObj } from './controls.ts'
import { PictureCollection } from './pictures.ts'
import {
  NullRangeObj, SpanRangeObj, compareLocation, findInParagraph, intersect, packageOf, paragraphXml, readInlineOoxml, textRanges,
} from './ranges.ts'
import type { FieldObj } from './review.ts'
import { bookmarksOf, insertComment, insertField } from './review.ts'
import { TableObj, insertTableAt } from './tables.ts'

export { ControlCollection, ControlObj } from './controls.ts'
export { CellObj, NullTableObj, TableCollection, TableObj } from './tables.ts'

export class FontObj extends FakeObject {
  constructor(context: Ctx, private readonly target: () => FakeFont[]) {
    super(context)
  }
  private read<K extends keyof FakeFont>(key: K): FakeFont[K] {
    this.need(key === 'highlightColor' ? 'highlightColor' : key)
    const values = this.target().map((f) => f[key])
    const first = values[0] ?? null
    return (values.every((v) => v === first) ? first : null) as FakeFont[K]
  }
  private write<K extends keyof FakeFont>(key: K, value: FakeFont[K]): void {
    this.context.queue(() => this.target().forEach((f) => (f[key] = value)))
  }
  get name(): string | null { return this.read('name') }
  set name(v: string | null) { this.write('name', v) }
  get size(): number | null { return this.read('size') }
  set size(v: number | null) { this.write('size', v) }
  get color(): string | null { return this.read('color') }
  set color(v: string | null) { this.write('color', v) }
  get highlightColor(): string | null { return this.read('highlightColor') }
  set highlightColor(v: string | null) { this.write('highlightColor', v) }
  get bold(): boolean | null { return this.read('bold') }
  set bold(v: boolean | null) { this.write('bold', v) }
  get italic(): boolean | null { return this.read('italic') }
  set italic(v: boolean | null) { this.write('italic', v) }
  set(values: Partial<FakeFont>): void {
    this.context.queue(() => this.target().forEach((f) => Object.assign(f, values)))
  }
  /** Font.reset (WordApiDesktop 1.3): every direct character format goes, bold and italic included. */
  reset(): void {
    this.set({ name: null, size: null, color: null, highlightColor: null, bold: null, italic: null })
  }
}

/** Where a paragraph lives: the body or a header or footer part. */
export interface Container {
  model: FakeModel
  list: FakeParagraph[]
  controls: FakeControl[]
}

/** 'content' ranges (a paragraph without its mark, or part of one) take insertOoxml Replace inline. */
export type RangeKind = 'text' | 'content' | 'whole'

const liveRanges = new WeakMap<FakeParagraph, RangeObj[]>()

/** Keeps ranges on a paragraph in step with insertText edits made earlier in the same batch. */
export function trackRanges(p: FakeParagraph, ranges: RangeObj[]): void {
  liveRanges.set(p, [...(liveRanges.get(p) ?? []), ...ranges])
}

export class RangeObj extends FakeObject {
  /** Set on the object getSelection() returns, so intersections know what is selected. */
  selectionOf: FakeModel | null = null
  constructor(
    context: Ctx,
    readonly container: Container,
    readonly para: FakeParagraph | null,
    public start: number,
    public end: number,
    readonly inserted: { paragraphs: FakeParagraph[]; tables: FakeTable[]; controls: FakeControl[] } | null = null,
    readonly kind: RangeKind = 'text',
  ) {
    super(context)
  }
  get isNullObject(): boolean { return false }
  get text(): string {
    this.need('text')
    return this.para ? this.para.text.slice(this.start, this.end) : (this.inserted?.paragraphs.map((p) => p.text).join('\r') ?? '')
  }
  get font(): FontObj {
    return new FontObj(this.context, () => (this.para ? [this.para.font] : []))
  }
  set case(v: string) {
    this.context.queue(() => {
      const p = this.para
      if (!p) return
      const part = p.text.slice(this.start, this.end)
      const changed = v === 'Upper' ? part.toUpperCase() : v === 'Lower' ? part.toLowerCase() : part
      p.text = p.text.slice(0, this.start) + changed + p.text.slice(this.end)
    })
  }
  insertText(text: string, location: string): RangeObj {
    this.context.queue(() => {
      if (!this.para || location !== 'Replace') throw new Error(`fake insertText supports Replace on text ranges only (${location})`)
      const p = this.para
      const delta = text.length - (this.end - this.start)
      p.text = p.text.slice(0, this.start) + text + p.text.slice(this.end)
      for (const r of liveRanges.get(p) ?? []) {
        if (r !== this && r.start >= this.end) {
          r.start += delta
          r.end += delta
        }
      }
      this.end = this.start + text.length
    })
    return this
  }
  delete(): void {
    this.context.queue(() => {
      if (this.inserted) {
        const { model } = this.container
        const inserted = this.inserted
        const keep = <T>(list: T[], drop: T[]): void => {
          const kept = list.filter((x) => !drop.includes(x))
          list.splice(0, list.length, ...kept)
        }
        keep(model.paragraphs, inserted.paragraphs)
        keep(model.tables, inserted.tables)
        keep(model.controls, inserted.controls)
        // Real Word can leave an empty paragraph behind after deleting inserted content at the end.
        if (model.ooxmlQuirk) model.paragraphs.push(para(''))
        return
      }
      if (this.para) {
        this.para.text = this.para.text.slice(0, this.start) + this.para.text.slice(this.end)
      }
    })
  }
  select(mode = 'Select'): void {
    this.context.queue(() => {
      if (this.para && !this.selectionOf) this.container.model.selection = [this.para]
      this.container.model.selected.push(`range:${mode}:${this.para?.text.slice(this.start, this.end) ?? 'inserted'}`)
    })
  }
  insertOoxml(ooxml: string, location: string): RangeObj {
    const p = this.para
    if (!p) throw new Error('fake insertOoxml needs a paragraph range')
    if (location !== 'Replace' || this.kind !== 'content') return insertOoxmlAt(this.context, this.container, p, ooxml, location)
    // Inside a paragraph, Word puts the runs in place and keeps the paragraph mark.
    const range = new RangeObj(this.context, this.container, p, this.start, this.start, null, 'content')
    this.context.queue(() => {
      const { text, font } = readInlineOoxml(ooxml)
      this.container.model.inserts.push({ location: 'Inline', ooxml })
      p.text = p.text.slice(0, this.start) + text + p.text.slice(this.end)
      Object.assign(p.font, font)
      range.end = this.start + text.length
    })
    return range
  }
  getOoxml(): { value: string } {
    return clientResult(this.context, '', () => (this.para ? packageOf(paragraphXml(this.container.model, this.para, this.start, this.end)) : ''))
  }
  getRange(location = 'Whole'): RangeObj {
    const at = location === 'Start' ? this.start : location === 'End' || location === 'After' ? this.end : null
    return new RangeObj(this.context, this.container, this.para, at ?? this.start, at ?? this.end, null, this.kind)
  }
  getTextRanges(marks: string[], trim = false): RangeCollection {
    return new RangeCollection(this.context, () => textRanges(this, marks, trim))
  }
  intersectWithOrNullObject(other: RangeObj): RangeObj | NullRangeObj { return intersect(this, other) }
  expandTo(other: RangeObj): SpanRangeObj {
    const list = this.container.list
    const from = this.para ? list.indexOf(this.para) : 0
    const to = other.para ? list.indexOf(other.para) : from
    return new SpanRangeObj(this.context, this.container, list.slice(Math.min(from, to), Math.max(from, to) + 1))
  }
  compareLocationWith(other: RangeObj): { value: string } {
    return clientResult(this.context, '', () => compareLocation(this, other))
  }
  search(needle: string, options: { matchCase?: boolean } = {}): RangeCollection {
    const p = this.para
    return new RangeCollection(this.context, () =>
      p ? findInParagraph(this.context, this.container, p, needle, Boolean(options.matchCase), this.start, this.end, this.kind === 'whole') : [],
    )
  }
  insertBookmark(name: string): void {
    const p = this.para
    this.context.queue(() => {
      if (!p || !/^[A-Za-z][A-Za-z0-9_]{0,39}$/.test(name)) throw Object.assign(new Error(`Bad bookmark ${name}`), { code: 'InvalidArgument' })
      bookmarksOf(this.container.model).set(name, { paragraph: p, start: this.start, end: this.end })
    })
  }
  getBookmarks(_includeHidden?: boolean, _includeAdjacent?: boolean): { value: string[] } {
    return clientResult(this.context, [], () =>
      [...bookmarksOf(this.container.model)]
        .filter(([, b]) => b.paragraph === this.para && b.start < this.end && b.end > this.start)
        .map(([name]) => name),
    )
  }
  insertComment(text: string): unknown { return insertComment(this.context, this.container.model, text) }
  insertField(_location: string, type: string, text: string): FieldObj { return insertField(this.context, this.container.model, type, text) }
}

export class RangeCollection extends FakeCollection<FakeObject> {
  constructor(context: Ctx, private readonly build: () => FakeObject[]) {
    super(context)
  }
  protected make(): FakeObject[] {
    return this.build()
  }
}

/** Inserts an OOXML package relative to a paragraph, as Word does, and returns the inserted range. */
export function insertOoxmlAt(context: Ctx, container: Container, anchor: FakeParagraph | null, ooxml: string, location: string): RangeObj {
  const inserted = { paragraphs: [] as FakeParagraph[], tables: [] as FakeTable[], controls: [] as FakeControl[] }
  context.queue(() => {
    const { model } = container
    const content = ingestOoxml(ooxml)
    model.inserts.push({ location, ooxml })
    content.styleNames.forEach((s) => model.styles.add(s))
    inserted.paragraphs.push(...content.paragraphs)
    inserted.tables.push(...content.tables)
    inserted.controls.push(...content.controls)
    const list = model.paragraphs
    let at = anchor ? list.indexOf(anchor) : list.length
    if (location === 'End' || at < 0) at = list.length
    else if (location === 'After') at += 1
    else if (location === 'Replace' && anchor) list.splice(at, 1)
    list.splice(at, 0, ...content.paragraphs)
    model.tables.push(...content.tables)
    model.controls.push(...content.controls)
  })
  return new RangeObj(context, container, null, 0, 0, inserted)
}

export class ParagraphObj extends FakeObject {
  constructor(context: Ctx, readonly container: Container, readonly data: FakeParagraph) {
    super(context)
  }
  get isNullObject(): boolean { return false }
  get text(): string { this.need('text'); return this.data.text }
  get style(): string { this.need('style'); return this.data.style }
  set style(v: string) {
    this.context.queue(() => {
      if (!this.container.model.styles.has(v)) throw Object.assign(new Error(`Style ${v} not found`), { code: 'ItemNotFound' })
      this.data.style = v
      this.data.styleBuiltIn = 'Other'
    })
  }
  get styleBuiltIn(): string { this.need('styleBuiltIn'); return this.data.styleBuiltIn }
  set styleBuiltIn(v: string) {
    this.context.queue(() => {
      this.data.styleBuiltIn = v
      const h = /^Heading([1-9])$/.exec(v)
      this.data.style = h ? `heading ${h[1]}` : v
    })
  }
  get isListItem(): boolean { this.need('isListItem'); return this.data.isListItem }
  get tableNestingLevel(): number { this.need('tableNestingLevel'); return this.data.tableNestingLevel }
  get isLastParagraph(): boolean { this.need('isLastParagraph'); return this.container.list[this.container.list.length - 1] === this.data }
  get alignment(): string { this.need('alignment'); return this.data.alignment ?? 'Left' }
  set alignment(v: string) { this.context.queue(() => (this.data.alignment = v)) }
  set spaceAfter(_v: number) {}
  set spaceBefore(_v: number) {}
  get font(): FontObj {
    const f = new FontObj(this.context, () => [this.data.font])
    if (this.loaded.has('font')) f.markLoaded('*')
    return f
  }
  protected override onLoaded(head: string, rest: string): void {
    if (head === 'font' && rest) this.loaded.add('font')
  }
  get inlinePictures(): PictureCollection {
    const model = this.container.model
    return new PictureCollection(this.context, this.container, () => model.pictures.filter((p) => p.paragraph === this.data))
  }
  getRange(location = 'Whole'): RangeObj {
    const n = this.data.text.length
    const at = location === 'Start' ? 0 : location === 'End' || location === 'After' ? n : null
    const kind: RangeKind = location === 'Content' ? 'content' : location === 'Whole' ? 'whole' : 'text'
    return new RangeObj(this.context, this.container, this.data, at ?? 0, at ?? n, null, kind)
  }
  private neighbour(offset: 1 | -1): NeighbourObj {
    const list = this.container.list
    const i = list.indexOf(this.data)
    return new NeighbourObj(this.context, this.container, i >= 0 ? (list[i + offset] ?? null) : null)
  }
  getNextOrNullObject(): NeighbourObj { return this.neighbour(1) }
  getPreviousOrNullObject(): NeighbourObj { return this.neighbour(-1) }
  search(needle: string, options: { matchCase?: boolean } = {}): RangeCollection {
    return new RangeCollection(this.context, () =>
      findInParagraph(this.context, this.container, this.data, needle, Boolean(options.matchCase), 0, this.data.text.length, true, true),
    )
  }
  getOoxml(): { value: string } {
    return clientResult(this.context, '', () => packageOf(paragraphXml(this.container.model, this.data)))
  }
  insertOoxml(ooxml: string, location: string): RangeObj {
    return insertOoxmlAt(this.context, this.container, this.data, ooxml, location)
  }
  insertParagraph(text: string, location: string): ParagraphObj {
    const p = para(text)
    this.context.queue(() => {
      const list = this.container.list
      list.splice(list.indexOf(this.data) + (location === 'Before' ? 0 : 1), 0, p)
    })
    return new ParagraphObj(this.context, this.container, p)
  }
  insertTable(_rows: number, _columns: number, location: string, values: string[][]): TableObj {
    return insertTableAt(this.context, this.container, this.data, location, values)
  }
  insertContentControl(): ControlObj {
    const control: FakeControl = { tag: '', title: '', paragraphs: [this.data] }
    this.context.queue(() => this.container.controls.push(control))
    return new ControlObj(this.context, this.container, control)
  }
  get parentTable(): TableObj {
    const table = this.data.table
    if (!table) throw new Error('paragraph is not in a table')
    return new TableObj(this.context, this.container, table)
  }
  clear(): void { this.context.queue(() => (this.data.text = '')) }
  delete(): void {
    this.context.queue(() => {
      const i = this.container.list.indexOf(this.data)
      if (i >= 0) this.container.list.splice(i, 1)
    })
  }
  select(_mode?: string): void {
    this.context.queue(() => {
      this.container.model.selection = [this.data]
      this.container.model.selected.push(`paragraph:${this.data.text}`)
    })
  }
}

export class ParagraphCollection extends FakeCollection<ParagraphObj> {
  constructor(context: Ctx, private readonly container: Container, private readonly source: () => FakeParagraph[]) {
    super(context)
  }
  protected make(): ParagraphObj[] {
    return this.source().map((p) => new ParagraphObj(this.context, this.container, p))
  }
  getFirst(): ParagraphObj {
    const first = this.source()[0]
    if (!first) throw new Error('no paragraphs')
    return new ParagraphObj(this.context, this.container, first)
  }
  getLast(): ParagraphObj {
    const all = this.source()
    const last = all[all.length - 1]
    if (!last) throw new Error('no paragraphs')
    return new ParagraphObj(this.context, this.container, last)
  }
}

export class NeighbourObj extends ParagraphObj {
  constructor(context: Ctx, container: Container, private readonly maybe: FakeParagraph | null) {
    super(context, container, maybe ?? para(''))
  }
  override get isNullObject(): boolean { return this.maybe === null }
}

export function partContainer(model: FakeModel, part: { paragraphs: FakeParagraph[]; controls: FakeControl[] }): Container {
  return { model, list: part.paragraphs, controls: part.controls }
}
