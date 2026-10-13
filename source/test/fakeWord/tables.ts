// Fake Office.js tables: rows, cells, layout calls, adding and deleting rows, and inserting a table after a paragraph.

import { Ctx, FakeCollection, FakeObject } from './base.ts'
import type { FakeParagraph, FakeTable } from './model.ts'
import { para } from './model.ts'
import type { Container } from './objects.ts'
import { NeighbourObj, ParagraphObj } from './objects.ts'
import { PictureCollection } from './pictures.ts'

const invalid = (message: string): Error => Object.assign(new Error(message), { code: 'InvalidArgument' })

const isUniform = (t: FakeTable): boolean => t.rows.every((r) => r.length === (t.rows[0]?.length ?? 0))

export class TableObj extends FakeObject {
  constructor(context: Ctx, private readonly container: Container, readonly data: FakeTable) {
    super(context)
  }
  get isNullObject(): boolean { return false }
  get values(): string[][] { this.need('values'); return this.data.rows.map((r) => [...r]) }
  set values(v: string[][]) { this.context.queue(() => (this.data.rows = v.map((r) => [...r]))) }
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
  get isUniform(): boolean { this.need('isUniform'); return isUniform(this.data) }
  get alignment(): string { this.need('alignment'); return this.data.alignment ?? 'Left' }
  set alignment(v: string) { this.context.queue(() => (this.data.alignment = v)) }
  private act(action: string, uniformOnly = false): void {
    this.context.queue(() => {
      if (uniformOnly && !isUniform(this.data)) throw invalid('The table is not uniform.')
      this.data.actions = [...(this.data.actions ?? []), action]
    })
  }
  distributeColumns(): void { this.act('distributeColumns', true) }
  autoFitWindow(): void { this.act('autoFitWindow') }
  autoFitBehavior(behavior: string): void { this.act(`autoFitBehavior:${behavior}`) }
  get rows(): TableRowCollection {
    return new TableRowCollection(this.context, this.data, () => this.data.rows.map((_, i) => i))
  }
  addRows(location: string, count: number, values?: string[][]): TableRowCollection {
    const start = location === 'End' ? null : 0
    const added: number[] = []
    this.context.queue(() => {
      const width = this.data.rows[this.data.rows.length - 1]?.length ?? 0
      const fresh = Array.from({ length: count }, (_, i) => Array.from({ length: width }, (_, c) => values?.[i]?.[c] ?? ''))
      const at = start ?? this.data.rows.length
      this.data.rows.splice(at, 0, ...fresh)
      added.push(...fresh.map((_, i) => at + i))
    })
    return new TableRowCollection(this.context, this.data, () => added)
  }
  deleteRows(index: number, count = 1): void {
    this.context.queue(() => {
      if (index < 0 || index + count > this.data.rows.length) throw invalid('No such rows.')
      this.data.rows.splice(index, count)
    })
  }
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
  getCellOrNullObject(row: number, col: number): CellObj { return new CellObj(this.context, this.data, row, col) }
  /** Like Office.js, the cell is looked up when the batch runs, so a row added earlier in the batch can be used. */
  getCell(row: number, col: number): CellObj {
    return new CellObj(this.context, this.data, row, col)
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
  delete(): void {
    this.context.queue(() => {
      const { model, list } = this.container
      const i = model.tables.indexOf(this.data)
      if (i >= 0) model.tables.splice(i, 1)
      const kept = list.filter((p) => p.table !== this.data)
      list.splice(0, list.length, ...kept)
      if (model.selectionTable === this.data) delete model.selectionTable
    })
  }
  select(_mode?: string): void {
    this.context.queue(() => this.container.model.selected.push(`table:${this.data.style}`))
  }
}

/** Inserts a table after (or before) a paragraph, with one paragraph per cell so neighbours and positions work. */
export function insertTableAt(context: Ctx, container: Container, anchor: FakeParagraph, location: string, values: string[][]): TableObj {
  const data: FakeTable = { style: 'Table Grid', rows: values.map((r) => [...r]) }
  context.queue(() => {
    const { model, list } = container
    const at = list.indexOf(anchor) + (location === 'Before' ? 0 : 1)
    const cells = data.rows.flatMap((r) => r.map((text) => para(text, 'Normal', { tableNestingLevel: 1, table: data })))
    list.splice(at, 0, ...cells)
    // Keep model.tables in document order: after every table whose cells come earlier.
    const before = model.tables.filter((t) => list.findIndex((p) => p.table === t) < at && list.some((p) => p.table === t))
    const last = before[before.length - 1]
    model.tables.splice(last ? model.tables.indexOf(last) + 1 : 0, 0, data)
  })
  return new TableObj(context, container, data)
}

export class CellObj extends FakeObject {
  constructor(context: Ctx, private readonly table: FakeTable, readonly rowIndex: number, readonly cellIndex: number) {
    super(context)
  }
  private get key(): string { return `${this.rowIndex},${this.cellIndex}` }
  get isNullObject(): boolean { return this.table.rows[this.rowIndex]?.[this.cellIndex] === undefined }
  get value(): string { this.need('value'); return this.table.rows[this.rowIndex]?.[this.cellIndex] ?? '' }
  set value(v: string) {
    this.context.queue(() => {
      const row = this.table.rows[this.rowIndex]
      if (!row || row[this.cellIndex] === undefined) throw invalid('The cell is a null object')
      row[this.cellIndex] = v
    })
  }
  set horizontalAlignment(v: string) {
    this.context.queue(() => {
      if (this.isNullObject) throw invalid('The cell is a null object')
      this.table.alignments = { ...this.table.alignments, [this.key]: v }
    })
  }
  set columnWidth(v: number) {
    this.context.queue(() => (this.table.widths = { ...this.table.widths, [this.key]: v }))
  }
  get body(): { inlinePictures: PictureCollection } {
    const count = this.table.pictures?.[this.key] ?? 0
    const pictures = Array.from({ length: count }, () => ({ altTextTitle: '', altTextDescription: '' }))
    return { inlinePictures: new PictureCollection(this.context, null, () => pictures) }
  }
}

export class CellCollection extends FakeCollection<CellObj> {
  constructor(context: Ctx, private readonly table: FakeTable, private readonly row: number) {
    super(context)
  }
  protected make(): CellObj[] {
    return (this.table.rows[this.row] ?? []).map((_, c) => new CellObj(this.context, this.table, this.row, c))
  }
}

/** A table row. Its index is looked up when the batch runs, so rows added in the same batch can be used. */
export class TableRowObj extends FakeObject {
  constructor(context: Ctx, private readonly table: FakeTable, private readonly at: () => number) {
    super(context)
  }
  get index(): number { return this.at() }
  get cellCount(): number { this.need('cellCount'); return this.table.rows[this.at()]?.length ?? 0 }
  get cells(): CellCollection { return new CellCollection(this.context, this.table, this.at()) }
  get font(): { bold: boolean } {
    const queue = (fn: () => void): void => this.context.queue(fn)
    const { table, at } = this
    return {
      get bold(): boolean { return (table.boldRows ?? []).includes(at()) },
      set bold(v: boolean) {
        queue(() => {
          const rest = (table.boldRows ?? []).filter((r) => r !== at())
          table.boldRows = v ? [...rest, at()] : rest
        })
      },
    }
  }
}

export class TableRowCollection extends FakeCollection<TableRowObj> {
  constructor(context: Ctx, private readonly table: FakeTable, private readonly indexes: () => number[]) {
    super(context)
  }
  protected make(): TableRowObj[] {
    return this.indexes().map((i) => new TableRowObj(this.context, this.table, () => i))
  }
  getFirst(): TableRowObj {
    return new TableRowObj(this.context, this.table, () => this.indexes()[0] ?? -1)
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
