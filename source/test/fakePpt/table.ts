// Fake tables: values, cell text, and cell formatting (PowerPointApi 1.9) recorded in the model's table.cells.

import { Ctx, FakeObject, FakeOfficeError } from './base.ts'
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
}
