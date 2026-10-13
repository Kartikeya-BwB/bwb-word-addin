// Fake tables: values, cell text, and cell formatting (PowerPointApi 1.9) recorded in the model's table.cells.

import { Ctx, FakeClientResult, FakeCollection, FakeObject, FakeOfficeError } from './base.ts'
import type { FakePresentation, FakeShape, FakeTable } from './model.ts'
import { ShapeObj } from './shapes.ts'
import { normaliseColour } from './text.ts'

class CellPartObj extends FakeObject {
  constructor(context: Ctx, protected readonly put: (values: Record<string, unknown>) => void) {
    super(context)
  }
}
class CellFillObj extends CellPartObj {
  setSolidColor(color: string): void { this.put({ color: normaliseColour(color) }) }
  clear(): void { this.put({ color: null }) }
}
class CellFontObj extends CellPartObj {
  set name(v: string) { this.put({ name: v }) }
  set size(v: number) { this.put({ size: v }) }
  set color(v: string) { this.put({ color: normaliseColour(v) }) }
  set bold(v: boolean) { this.put({ bold: v }) }
  set italic(v: boolean) { this.put({ italic: v }) }
}

class TableCellObj extends FakeObject {
  constructor(context: Ctx, private readonly table: () => FakeTable, private readonly row: number, private readonly column: number) {
    super(context)
  }
  protected override isNull(): boolean { return this.table().values[this.row]?.[this.column] === undefined }
  /** Cell formatting (PowerPointApi 1.9) is recorded in the table's cells, keyed "row,column". */
  private record(part: string, values: Record<string, unknown>): void {
    this.context.gate('1.9', `TableCell.${part}`)
    this.write(() => {
      const cell = (this.table().cells[`${this.row},${this.column}`] ??= {})
      cell[part] = { ...(cell[part] as Record<string, unknown> | undefined), ...values }
    })
  }
  get text(): string { this.need('text'); return this.table().values[this.row]?.[this.column] ?? '' }
  set text(v: string) {
    this.write(() => {
      const r = this.table().values[this.row]
      if (r) r[this.column] = v
    })
  }
  get rowIndex(): number { this.need('rowIndex'); return this.row }
  get columnIndex(): number { this.need('columnIndex'); return this.column }
  get fill(): CellFillObj { return new CellFillObj(this.context, (v) => this.record('fill', v)) }
  get font(): CellFontObj { return new CellFontObj(this.context, (v) => this.record('font', v)) }
  set horizontalAlignment(v: string) { this.record('paragraph', { horizontalAlignment: v }) }
  set verticalAlignment(v: string) { this.record('paragraph', { verticalAlignment: v }) }
  /** Borders (1.9), recorded as cells["row,column"].borders.top and so on. */
  get borders(): BordersObj {
    return new BordersObj(this.context, (side) => new BorderObj(this.context, (values) => {
      this.context.gate('1.9', 'TableCell.borders')
      this.write(() => {
        const cell = (this.table().cells[`${this.row},${this.column}`] ??= {})
        const borders = (cell.borders ??= {}) as Record<string, Record<string, unknown>>
        borders[side] = { ...borders[side], ...values }
      })
    }))
  }
}

class BorderObj extends CellPartObj {
  set color(v: string) { this.put({ color: normaliseColour(v) }) }
  set weight(v: number) { this.put({ weight: v }) }
  set dashStyle(v: string) { this.put({ dashStyle: v }) }
  set transparency(v: number) { this.put({ transparency: v }) }
}

class BordersObj extends FakeObject {
  constructor(context: Ctx, private readonly side: (name: string) => BorderObj) {
    super(context)
  }
  get top(): BorderObj { return this.side('top') }
  get bottom(): BorderObj { return this.side('bottom') }
  get left(): BorderObj { return this.side('left') }
  get right(): BorderObj { return this.side('right') }
  get diagonalDown(): BorderObj { return this.side('diagonalDown') }
  get diagonalUp(): BorderObj { return this.side('diagonalUp') }
}

/** Moves cell formatting when rows or columns come and go: `move` gives the new [row, column], or null to drop it. */
function shiftCells(t: FakeTable, move: (row: number, column: number) => [number, number] | null): void {
  const out: FakeTable['cells'] = {}
  for (const [key, value] of Object.entries(t.cells)) {
    const [r = 0, c = 0] = key.split(',').map(Number)
    const to = move(r, c)
    if (to) out[`${to[0]},${to[1]}`] = value
  }
  t.cells = out
}

const insertAt = (index: number | null | undefined, length: number): number =>
  index === undefined || index === null || index < 0 || index > length ? length : index

/** Column widths in points: equal shares of the table's width until set. */
function widthsOf(shape: FakeShape, t: FakeTable): number[] {
  const n = t.values[0]?.length ?? 0
  if (t.columnWidths?.length !== n) t.columnWidths = Array.from({ length: n }, () => shape.width / Math.max(1, n))
  return t.columnWidths
}

class TableRowObj extends FakeObject {
  constructor(context: Ctx, private readonly table: () => FakeTable, readonly index: number) {
    super(context)
  }
  get rowIndex(): number { this.need('rowIndex'); return this.index }
  get height(): number { this.need('height'); return this.table().rowHeights?.[this.index] ?? 30 }
  set height(v: number) {
    this.write(() => {
      const t = this.table()
      t.rowHeights ??= t.values.map(() => 30)
      t.rowHeights[this.index] = v
    })
  }
  get currentHeight(): number { this.need('currentHeight'); return this.table().rowHeights?.[this.index] ?? 30 }
  delete(): void { this.write(() => removeRow(this.table(), this.index)) }
}

function removeRow(t: FakeTable, index: number): void {
  t.values.splice(index, 1)
  t.rowHeights?.splice(index, 1)
  shiftCells(t, (r, c) => (r === index ? null : [r > index ? r - 1 : r, c]))
}

class TableColumnObj extends FakeObject {
  constructor(context: Ctx, private readonly shape: () => FakeShape, private readonly table: () => FakeTable, readonly index: number) {
    super(context)
  }
  get columnIndex(): number { this.need('columnIndex'); return this.index }
  get width(): number { this.need('width'); return widthsOf(this.shape(), this.table())[this.index] ?? 0 }
  set width(v: number) {
    this.write(() => {
      const widths = widthsOf(this.shape(), this.table())
      if (this.index >= widths.length) throw new FakeOfficeError('InvalidArgument', `No column ${this.index}`)
      widths[this.index] = v
      this.shape().width = widths.reduce((a, b) => a + b, 0)
    })
  }
  delete(): void { this.write(() => removeColumn(this.shape(), this.table(), this.index)) }
}

function removeColumn(shape: FakeShape, t: FakeTable, index: number): void {
  const widths = widthsOf(shape, t)
  t.values.forEach((row) => row.splice(index, 1))
  widths.splice(index, 1)
  shape.width = widths.reduce((a, b) => a + b, 0)
  shiftCells(t, (r, c) => (c === index ? null : [r, c > index ? c - 1 : c]))
}

class TableRowCollectionObj extends FakeCollection<TableRowObj> {
  private readonly keys = new Map<number, { index: number }>()
  constructor(context: Ctx, private readonly table: () => FakeTable) {
    super(context)
  }
  protected source(): object[] {
    return this.table().values.map((_, index) => {
      const key = this.keys.get(index) ?? { index }
      this.keys.set(index, key)
      return key
    })
  }
  protected wrap(data: object): TableRowObj { return new TableRowObj(this.context, this.table, (data as { index: number }).index) }
  getItemAt(index: number): TableRowObj { return new TableRowObj(this.context, this.table, index) }
  getCount(): FakeClientResult<number> { return new FakeClientResult(this.context, () => this.table().values.length) }
  /** Adds empty rows at `index` (at the end when it is missing or out of range), shifting cell formatting down. */
  add(index?: number | null, count?: number): void {
    this.write(() => {
      const t = this.table()
      const n = count || 1
      const at = insertAt(index, t.values.length)
      const columns = t.values[0]?.length ?? 0
      t.values.splice(at, 0, ...Array.from({ length: n }, () => Array.from({ length: columns }, () => '')))
      t.rowHeights?.splice(at, 0, ...Array.from({ length: n }, () => 30))
      shiftCells(t, (r, c) => [r >= at ? r + n : r, c])
    })
  }
  deleteRows(rows: TableRowObj[]): void {
    this.write(() => [...rows].sort((a, b) => b.index - a.index).forEach((r) => removeRow(this.table(), r.index)))
  }
}

class TableColumnCollectionObj extends FakeCollection<TableColumnObj> {
  private readonly keys = new Map<number, { index: number }>()
  constructor(context: Ctx, private readonly shape: () => FakeShape, private readonly table: () => FakeTable) {
    super(context)
  }
  protected source(): object[] {
    return (this.table().values[0] ?? []).map((_, index) => {
      const key = this.keys.get(index) ?? { index }
      this.keys.set(index, key)
      return key
    })
  }
  protected wrap(data: object): TableColumnObj { return new TableColumnObj(this.context, this.shape, this.table, (data as { index: number }).index) }
  getItemAt(index: number): TableColumnObj { return new TableColumnObj(this.context, this.shape, this.table, index) }
  getCount(): FakeClientResult<number> { return new FakeClientResult(this.context, () => this.table().values[0]?.length ?? 0) }
  /** Adds empty columns at `index`, each as wide as the average column; the table grows to fit. */
  add(index?: number | null, count?: number): void {
    this.write(() => {
      const shape = this.shape()
      const t = this.table()
      const widths = widthsOf(shape, t)
      const n = count || 1
      const at = insertAt(index, widths.length)
      const average = widths.length ? widths.reduce((a, b) => a + b, 0) / widths.length : 100
      t.values.forEach((row) => row.splice(at, 0, ...Array.from({ length: n }, () => '')))
      widths.splice(at, 0, ...Array.from({ length: n }, () => average))
      shape.width = widths.reduce((a, b) => a + b, 0)
      shiftCells(t, (r, c) => [r, c >= at ? c + n : c])
    })
  }
  deleteColumns(columns: TableColumnObj[]): void {
    this.write(() => [...columns].sort((a, b) => b.index - a.index).forEach((c) => removeColumn(this.shape(), this.table(), c.index)))
  }
}

const STYLE_DEFAULTS: Record<string, unknown> = {
  style: 'ThemedStyle1Accent1', areRowsBanded: true, areColumnsBanded: false, isFirstRowHighlighted: true, isLastRowHighlighted: false,
  isFirstColumnHighlighted: false, isLastColumnHighlighted: false,
}

/** Table style settings (1.9): reads give PowerPoint's defaults until set; writes land in table.styleSettings. */
class TableStyleSettingsObj extends FakeObject {
  constructor(context: Ctx, private readonly table: () => FakeTable) {
    super(context)
  }
  private get<T>(key: string): T {
    this.need(key)
    return (this.table().styleSettings?.[key] ?? STYLE_DEFAULTS[key]) as T
  }
  private set(key: string, value: unknown): void {
    this.write(() => {
      const t = this.table()
      t.styleSettings = { ...t.styleSettings, [key]: value }
    })
  }
  get style(): string { return this.get('style') }
  set style(v: string) { this.set('style', v) }
  get areRowsBanded(): boolean { return this.get('areRowsBanded') }
  set areRowsBanded(v: boolean) { this.set('areRowsBanded', v) }
  get areColumnsBanded(): boolean { return this.get('areColumnsBanded') }
  set areColumnsBanded(v: boolean) { this.set('areColumnsBanded', v) }
  get isFirstRowHighlighted(): boolean { return this.get('isFirstRowHighlighted') }
  set isFirstRowHighlighted(v: boolean) { this.set('isFirstRowHighlighted', v) }
  get isLastRowHighlighted(): boolean { return this.get('isLastRowHighlighted') }
  set isLastRowHighlighted(v: boolean) { this.set('isLastRowHighlighted', v) }
  get isFirstColumnHighlighted(): boolean { return this.get('isFirstColumnHighlighted') }
  set isFirstColumnHighlighted(v: boolean) { this.set('isFirstColumnHighlighted', v) }
  get isLastColumnHighlighted(): boolean { return this.get('isLastColumnHighlighted') }
  set isLastColumnHighlighted(v: boolean) { this.set('isLastColumnHighlighted', v) }
}

export class TableObj extends FakeObject {
  constructor(context: Ctx, private readonly model: FakePresentation, private readonly target: () => FakeShape) {
    super(context)
  }
  private data(): FakeTable {
    const t = this.target().table
    if (!t) throw new FakeOfficeError('InvalidArgument', `Shape ${this.target().id} is not a table`)
    return t
  }
  protected override checkLoad(): void { this.data() }
  get values(): string[][] { this.need('values'); return this.data().values.map((r) => [...r]) }
  get rowCount(): number { this.need('rowCount'); return this.data().values.length }
  get columnCount(): number { this.need('columnCount'); return this.data().values[0]?.length ?? 0 }
  getCellOrNullObject(row: number, column: number): TableCellObj { return new TableCellObj(this.context, () => this.data(), row, column) }
  getShape(): ShapeObj { return new ShapeObj(this.context, this.model, this.target()) }
  get rows(): TableRowCollectionObj {
    this.context.gate('1.9', 'Table.rows')
    return new TableRowCollectionObj(this.context, () => this.data())
  }
  get columns(): TableColumnCollectionObj {
    this.context.gate('1.9', 'Table.columns')
    return new TableColumnCollectionObj(this.context, () => this.target(), () => this.data())
  }
  get styleSettings(): TableStyleSettingsObj {
    this.context.gate('1.9', 'Table.styleSettings')
    return new TableStyleSettingsObj(this.context, () => this.data())
  }
}
