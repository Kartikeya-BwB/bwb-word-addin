// Fake Office.js objects for paragraphs, ranges, tables, content controls and header or footer bodies.

import { Ctx, FakeCollection, FakeObject } from './base.ts'
import type { FakeControl, FakeFont, FakeModel, FakeParagraph, FakePart, FakeTable } from './model.ts'
import { ingestOoxml, para } from './model.ts'

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
}

/** Where a paragraph lives: the body or a header or footer part. */
export interface Container {
  model: FakeModel
  list: FakeParagraph[]
  controls: FakeControl[]
}

export class RangeObj extends FakeObject {
  constructor(
    context: Ctx,
    readonly container: Container,
    readonly para: FakeParagraph | null,
    public start: number,
    public end: number,
    readonly inserted: { paragraphs: FakeParagraph[]; tables: FakeTable[]; controls: FakeControl[] } | null = null,
  ) {
    super(context)
  }
  get text(): string {
    this.need('text')
    return this.para ? this.para.text.slice(this.start, this.end) : (this.inserted?.paragraphs.map((p) => p.text).join('\r') ?? '')
  }
  get font(): FontObj {
    return new FontObj(this.context, () => (this.para ? [this.para.font] : []))
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
    this.context.queue(() => this.container.model.selected.push(`range:${mode}:${this.para?.text.slice(this.start, this.end) ?? 'inserted'}`))
  }
  insertOoxml(ooxml: string, location: string): RangeObj {
    if (!this.para) throw new Error('fake insertOoxml needs a paragraph range')
    return insertOoxmlAt(this.context, this.container, this.para, ooxml, location)
  }
}

const liveRanges = new WeakMap<FakeParagraph, RangeObj[]>()

export class RangeCollection extends FakeCollection<RangeObj> {
  constructor(context: Ctx, private readonly build: () => RangeObj[]) {
    super(context)
  }
  protected make(): RangeObj[] {
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
  getRange(_location = 'Whole'): RangeObj {
    return new RangeObj(this.context, this.container, this.data, 0, this.data.text.length)
  }
  search(needle: string, options: { matchCase?: boolean } = {}): RangeCollection {
    const target = needle.replace(/\^\^/g, '^')
    return new RangeCollection(this.context, () => {
      const hay = options.matchCase ? this.data.text : this.data.text.toLowerCase()
      const n = options.matchCase ? target : target.toLowerCase()
      const out: RangeObj[] = []
      for (let i = hay.indexOf(n); n && i !== -1; i = hay.indexOf(n, i + 1)) {
        out.push(new RangeObj(this.context, this.container, this.data, i, i + n.length))
      }
      liveRanges.set(this.data, [...(liveRanges.get(this.data) ?? []), ...out])
      return out
    })
  }
  insertOoxml(ooxml: string, location: string): RangeObj {
    return insertOoxmlAt(this.context, this.container, this.data, ooxml, location)
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
  delete(): void {
    this.context.queue(() => {
      const i = this.container.list.indexOf(this.data)
      if (i >= 0) this.container.list.splice(i, 1)
    })
  }
  select(): void {
    this.context.queue(() => this.container.model.selected.push(`paragraph:${this.data.text}`))
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
}

export class ControlObj extends FakeObject {
  constructor(context: Ctx, private readonly container: Container, readonly data: FakeControl) {
    super(context)
  }
  get tag(): string { this.need('tag'); return this.data.tag }
  set tag(v: string) { this.context.queue(() => (this.data.tag = v)) }
  get title(): string { this.need('title'); return this.data.title }
  set title(v: string) { this.context.queue(() => (this.data.title = v)) }
  set appearance(v: string) { this.context.queue(() => (this.data.appearance = v)) }
  get text(): string { this.need('text'); return this.data.paragraphs.map((p) => p.text).join('\r') }
  delete(keepContent: boolean): void {
    this.context.queue(() => {
      const i = this.container.controls.indexOf(this.data)
      if (i >= 0) this.container.controls.splice(i, 1)
      if (!keepContent) {
        const drop = new Set(this.data.paragraphs)
        const kept = this.container.list.filter((p) => !drop.has(p))
        this.container.list.splice(0, this.container.list.length, ...kept)
      }
    })
  }
  select(): void {
    this.context.queue(() => this.container.model.selected.push(`control:${this.data.tag}`))
  }
}

export class ControlCollection extends FakeCollection<ControlObj> {
  constructor(context: Ctx, private readonly container: Container, private readonly filter?: (c: FakeControl) => boolean) {
    super(context)
  }
  protected make(): ControlObj[] {
    return this.container.controls.filter((c) => !this.filter || this.filter(c)).map((c) => new ControlObj(this.context, this.container, c))
  }
  getByTag(tag: string): ControlCollection {
    return new ControlCollection(this.context, this.container, (c) => c.tag === tag)
  }
}

export class TableObj extends FakeObject {
  constructor(context: Ctx, private readonly container: Container, readonly data: FakeTable) {
    super(context)
  }
  get isNullObject(): boolean { return false }
  get values(): string[][] { this.need('values'); return this.data.rows.map((r) => [...r]) }
  get headerRowCount(): number { this.need('headerRowCount'); return this.data.headerRowCount ?? 0 }
  set headerRowCount(v: number) { this.context.queue(() => (this.data.headerRowCount = v)) }
  private option(name: string, v: boolean): void {
    this.context.queue(() => (this.data.options = { ...this.data.options, [name]: v }))
  }
  set styleFirstColumn(v: boolean) { this.option('styleFirstColumn', v) }
  set styleTotalRow(v: boolean) { this.option('styleTotalRow', v) }
  set styleBandedRows(v: boolean) { this.option('styleBandedRows', v) }
  set styleBandedColumns(v: boolean) { this.option('styleBandedColumns', v) }
  get style(): string { this.need('style'); return this.data.style }
  set style(v: string) {
    this.context.queue(() => {
      if (!this.container.model.styles.has(v)) throw Object.assign(new Error(`Style ${v} not found`), { code: 'ItemNotFound' })
      this.data.style = v
    })
  }
  get rowCount(): number { this.need('rowCount'); return this.data.rows.length }
  private cells(): FakeParagraph[] {
    return this.container.list.filter((p) => p.table === this.data)
  }
  private neighbour(offset: 1 | -1): NeighbourObj {
    const cells = this.cells()
    const list = this.container.list
    const edge = offset === 1 ? cells[cells.length - 1] : cells[0]
    const i = edge ? list.indexOf(edge) + offset : -1
    return new NeighbourObj(this.context, this.container, i >= 0 ? (list[i] ?? null) : null)
  }
  getParagraphAfterOrNullObject(): NeighbourObj { return this.neighbour(1) }
  getParagraphBeforeOrNullObject(): NeighbourObj { return this.neighbour(-1) }
  getCellOrNullObject(row: number, col: number): CellObj {
    return new CellObj(this.context, this.data.rows[row]?.[col] ?? null, (v) => {
      this.data.alignments = { ...this.data.alignments, [`${row},${col}`]: v }
    })
  }
  insertParagraph(text: string, location: string): ParagraphObj {
    const p = para(text)
    this.context.queue(() => {
      const cells = this.cells()
      const last = cells[cells.length - 1]
      const i = last ? this.container.list.indexOf(last) + 1 : this.container.list.length
      this.container.list.splice(location === 'After' ? i : 0, 0, p)
    })
    return new ParagraphObj(this.context, this.container, p)
  }
  select(): void {
    this.context.queue(() => this.container.model.selected.push(`table:${this.data.style}`))
  }
}

export class NeighbourObj extends ParagraphObj {
  constructor(context: Ctx, container: Container, private readonly maybe: FakeParagraph | null) {
    super(context, container, maybe ?? para(''))
  }
  override get isNullObject(): boolean { return this.maybe === null }
}

export class CellObj extends FakeObject {
  constructor(context: Ctx, private readonly text: string | null, private readonly align: (v: string) => void = () => undefined) {
    super(context)
  }
  get isNullObject(): boolean { return this.text === null }
  get value(): string { this.need('value'); return this.text ?? '' }
  set horizontalAlignment(v: string) {
    this.context.queue(() => {
      if (this.text === null) throw Object.assign(new Error('The cell is a null object'), { code: 'InvalidArgument' })
      this.align(v)
    })
  }
}

/** What getSelection().parentTableOrNullObject returns outside a table. */
export class NullTableObj extends FakeObject {
  get isNullObject(): boolean { return true }
}

export class TableCollection extends FakeCollection<TableObj> {
  constructor(context: Ctx, private readonly container: Container) {
    super(context)
  }
  protected make(): TableObj[] {
    return this.container.model.tables.map((t) => new TableObj(this.context, this.container, t))
  }
}

export function partContainer(model: FakeModel, part: FakePart): Container {
  return { model, list: part.paragraphs, controls: part.controls }
}
