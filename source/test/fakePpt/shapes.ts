// Fake shapes and what hangs off them: fill, outline, tags, placeholder format, groups and tables, plus the shape
// collections of a slide, a group and the selection.

import { Ctx, FakeClientResult, FakeCollection, FakeObject, FakeOfficeError } from './base.ts'
import type { FakePresentation, FakeShape, FakeSlide } from './model.ts'
import { group as groupData, line as lineData, shape as shapeData, tableShape, textBox as textBoxData } from './model.ts'
import type { Location } from './queries.ts'
import { boundsOf, locate, setBounds } from './queries.ts'
import { TextFrameObj, normaliseColour } from './text.ts'
import { TableObj } from './table.ts'
import { SlideObj } from './presentation.ts'

/** Shape properties newer than PowerPointApi 1.4. */
export const SHAPE_VERSIONS: Record<string, string> = {
  rotation: '1.10', visible: '1.10', isDecorative: '1.10', altTextTitle: '1.10', altTextDescription: '1.10', creationId: '1.10', adjustments: '1.10',
  zOrderPosition: '1.8', level: '1.8', placeholderFormat: '1.8', group: '1.8', parentGroup: '1.8',
}

const leaves = (s: FakeShape): FakeShape[] => (s.members?.length ? s.members.flatMap(leaves) : [s])

const notFound = (what: string): FakeOfficeError => new FakeOfficeError('ItemNotFound', `${what} was not found`)

class TagObj extends FakeObject {
  constructor(context: Ctx, private readonly map: () => Map<string, string>, readonly keyName: string, private readonly nullable: boolean) {
    super(context)
  }
  protected override isNull(): boolean { return this.nullable && !this.map().has(this.keyName) }
  protected override checkLoad(): void { if (!this.nullable && !this.map().has(this.keyName)) throw notFound(`Tag ${this.keyName}`) }
  get key(): string { this.need('key'); return this.keyName }
  get value(): string { this.need('value'); return this.map().get(this.keyName) ?? '' }
  set value(v: string) { this.write(() => this.map().set(this.keyName, v)) }
  delete(): void { this.write(() => this.map().delete(this.keyName)) }
}

/** Tags on a shape, a slide or the presentation. Keys are stored in upper case, as PowerPoint does. */
export class TagCollectionObj extends FakeCollection<TagObj> {
  private readonly keys = new Map<string, { key: string }>()
  constructor(context: Ctx, private readonly map: () => Map<string, string>) {
    super(context)
  }
  protected source(): object[] {
    return [...this.map().keys()].map((k) => {
      const found = this.keys.get(k) ?? { key: k }
      this.keys.set(k, found)
      return found
    })
  }
  protected wrap(data: object): TagObj { return new TagObj(this.context, this.map, (data as { key: string }).key, false) }
  add(key: string, value: string): void { this.write(() => this.map().set(key.toUpperCase(), value)) }
  delete(key: string): void { this.write(() => this.map().delete(key.toUpperCase())) }
  getItem(key: string): TagObj { return new TagObj(this.context, this.map, key.toUpperCase(), false) }
  getItemOrNullObject(key: string): TagObj { return new TagObj(this.context, this.map, key.toUpperCase(), true) }
  getItemAt(index: number): TagObj { return new TagObj(this.context, this.map, [...this.map().keys()][index] ?? '', false) }
  getCount(): FakeClientResult<number> { return new FakeClientResult(this.context, () => this.map().size) }
}

export class FillObj extends FakeObject {
  constructor(context: Ctx, private readonly target: () => FakeShape) {
    super(context)
  }
  private first(): FakeShape { return leaves(this.target())[0] ?? this.target() }
  private each(fn: (s: FakeShape) => void): void { this.write(() => leaves(this.target()).forEach(fn)) }
  get type(): string { this.need('type'); return this.first().fill.type }
  get foregroundColor(): string { this.need('foregroundColor'); return this.first().fill.foregroundColor }
  set foregroundColor(v: string) { this.each((s) => (s.fill = { ...s.fill, type: 'Solid', foregroundColor: normaliseColour(v) })) }
  get transparency(): number { this.need('transparency'); return this.first().fill.transparency }
  set transparency(v: number) { this.each((s) => (s.fill.transparency = v)) }
  setSolidColor(color: string): void { this.foregroundColor = color }
  clear(): void { this.each((s) => (s.fill.type = 'NoFill')) }
  setImage(_base64: string): void {
    this.context.gate('1.8', 'ShapeFill.setImage')
    this.each((s) => (s.fill.type = 'PictureAndTexture'))
  }
}

export class LineFormatObj extends FakeObject {
  constructor(context: Ctx, private readonly target: () => FakeShape) {
    super(context)
  }
  private first(): FakeShape { return leaves(this.target())[0] ?? this.target() }
  private put(fn: (s: FakeShape) => void): void { this.write(() => leaves(this.target()).forEach(fn)) }
  get visible(): boolean { this.need('visible'); return this.first().line.visible }
  set visible(v: boolean) { this.put((s) => (s.line.visible = v)) }
  get color(): string { this.need('color'); return this.first().line.color }
  set color(v: string) { this.put((s) => Object.assign(s.line, { color: normaliseColour(v), visible: true })) }
  get weight(): number { this.need('weight'); return this.first().line.weight }
  set weight(v: number) { this.put((s) => Object.assign(s.line, { weight: v, visible: true })) }
  get dashStyle(): string { this.need('dashStyle'); return this.first().line.dashStyle }
  set dashStyle(v: string) { this.put((s) => (s.line.dashStyle = v)) }
  get style(): string { this.need('style'); return this.first().line.style }
  set style(v: string) { this.put((s) => (s.line.style = v)) }
  get transparency(): number { this.need('transparency'); return this.first().line.transparency }
  set transparency(v: number) { this.put((s) => (s.line.transparency = v)) }
}

class PlaceholderFormatObj extends FakeObject {
  constructor(context: Ctx, private readonly target: () => FakeShape) {
    super(context)
  }
  protected override checkLoad(): void {
    if (this.target().type !== 'Placeholder') throw new FakeOfficeError('InvalidArgument', `Shape ${this.target().id} is not a placeholder`)
  }
  get type(): string { this.need('type'); return this.target().placeholderType ?? 'Unsupported' }
  get containedType(): string | null { this.need('containedType'); return this.target().textFrame ? null : 'Image' }
}

/** Puts `members` (top-level shapes of one list) into group `g`, where the top-most member was. */
function makeGroup(model: FakePresentation, list: FakeShape[], members: FakeShape[], g: FakeShape): void {
  if (members.length < 2 || members.some((m) => !list.includes(m))) {
    throw new FakeOfficeError('InvalidArgument', 'A group needs two or more shapes from the same slide or group')
  }
  const ordered = [...members].sort((a, b) => list.indexOf(a) - list.indexOf(b))
  const top = list.indexOf(ordered[ordered.length - 1] as FakeShape)
  const at = top - (ordered.length - 1)
  for (const m of ordered) list.splice(list.indexOf(m), 1)
  g.members = ordered
  list.splice(at, 0, g)
  const selected = new Set(model.selection.shapeIds)
  if (ordered.every((m) => selected.has(m.id))) model.selection = { ...model.selection, shapeIds: [g.id], textRange: null }
}

class ShapeGroupObj extends FakeObject {
  constructor(context: Ctx, private readonly model: FakePresentation, private readonly target: () => FakeShape) {
    super(context)
  }
  protected override checkLoad(): void {
    if (this.target().type !== 'Group') throw new FakeOfficeError('InvalidArgument', `Shape ${this.target().id} is not a group`)
  }
  get id(): string { this.need('id'); return this.target().id }
  get shape(): ShapeObj { return new ShapeObj(this.context, this.model, this.target()) }
  get shapes(): ShapeScopedCollectionObj { return new ShapeScopedCollectionObj(this.context, this.model, () => this.target().members ?? []) }
  ungroup(): void {
    this.write(() => {
      const g = this.target()
      const where = locate(this.model, g)
      if (g.type !== 'Group' || !where) throw new FakeOfficeError('InvalidArgument', `Shape ${g.id} is not a group on a slide`)
      const members = g.members ?? []
      where.list.splice(where.list.indexOf(g), 1, ...members)
      const sel = this.model.selection
      if (sel.shapeIds.includes(g.id)) this.model.selection = { ...sel, shapeIds: members.map((m) => m.id), textRange: null }
    })
  }
}

export class ShapeObj extends FakeObject {
  private fillObj: FillObj | undefined
  private lineObj: LineFormatObj | undefined
  private frameObj: TextFrameObj | undefined
  private tagsObj: TagCollectionObj | undefined
  private placeholderObj: PlaceholderFormatObj | undefined
  private groupObj: ShapeGroupObj | undefined

  /** `shape` null stands for an id that matched nothing: a null object when `nullable`, else a failing load. */
  constructor(context: Ctx, readonly model: FakePresentation, readonly shape: FakeShape | null, private readonly nullable = false) {
    super(context)
  }

  get data(): FakeShape {
    if (!this.shape) throw notFound('The shape')
    return this.shape
  }
  protected override isNull(): boolean { return this.nullable && !this.shape }
  protected override checkLoad(): void { if (!this.nullable && !this.shape) throw notFound('The shape') }
  protected override versionFor(path: string): string | undefined { return SHAPE_VERSIONS[path.split('/')[0] ?? ''] }

  private put(fn: (s: FakeShape) => void): void { this.write(() => fn(this.data)) }
  private where(): Location {
    const found = locate(this.model, this.data)
    if (!found) throw notFound(`Shape ${this.data.id}`)
    return found
  }
  private newer(version: string, prop: string): void {
    if (!this.context.supports(version)) throw new FakeOfficeError('ApiNotFound', `Shape.${prop} needs PowerPointApi ${version}`)
  }

  get id(): string { this.need('id'); return this.data.id }
  get name(): string { this.need('name'); return this.data.name }
  set name(v: string) { this.put((s) => (s.name = v)) }
  get type(): string { this.need('type'); return this.data.type }
  get left(): number { this.need('left'); return boundsOf(this.data).left }
  set left(v: number) { this.put((s) => setBounds(s, { left: v })) }
  get top(): number { this.need('top'); return boundsOf(this.data).top }
  set top(v: number) { this.put((s) => setBounds(s, { top: v })) }
  get width(): number { this.need('width'); return boundsOf(this.data).width }
  set width(v: number) { this.put((s) => setBounds(s, { width: v })) }
  get height(): number { this.need('height'); return boundsOf(this.data).height }
  set height(v: number) { this.put((s) => setBounds(s, { height: v })) }
  get rotation(): number { this.newer('1.10', 'rotation'); this.need('rotation'); return this.data.rotation }
  set rotation(v: number) { this.context.gate('1.10', 'Shape.rotation'); this.put((s) => (s.rotation = ((v % 360) + 360) % 360)) }
  get visible(): boolean { this.newer('1.10', 'visible'); this.need('visible'); return this.data.visible }
  set visible(v: boolean) { this.context.gate('1.10', 'Shape.visible'); this.put((s) => (s.visible = v)) }
  get isDecorative(): boolean { this.newer('1.10', 'isDecorative'); this.need('isDecorative'); return this.data.isDecorative }
  set isDecorative(v: boolean) { this.context.gate('1.10', 'Shape.isDecorative'); this.put((s) => (s.isDecorative = v)) }
  get altTextTitle(): string { this.newer('1.10', 'altTextTitle'); this.need('altTextTitle'); return this.data.altTextTitle }
  set altTextTitle(v: string) { this.context.gate('1.10', 'Shape.altTextTitle'); this.put((s) => (s.altTextTitle = v)) }
  get altTextDescription(): string { this.newer('1.10', 'altTextDescription'); this.need('altTextDescription'); return this.data.altTextDescription }
  set altTextDescription(v: string) { this.context.gate('1.10', 'Shape.altTextDescription'); this.put((s) => (s.altTextDescription = v)) }
  get zOrderPosition(): number { this.newer('1.8', 'zOrderPosition'); this.need('zOrderPosition'); return this.where().list.indexOf(this.data) }
  get level(): number {
    this.newer('1.8', 'level')
    this.need('level')
    let depth = 0
    for (let p = this.where().parent; p; p = locate(this.model, p)?.parent ?? null) depth += 1
    return depth
  }

  get fill(): FillObj { return (this.fillObj ??= new FillObj(this.context, () => this.data)) }
  get lineFormat(): LineFormatObj { return (this.lineObj ??= new LineFormatObj(this.context, () => this.data)) }
  get textFrame(): TextFrameObj { return (this.frameObj ??= new TextFrameObj(this.context, { model: this.model, shape: () => this.data })) }
  get tags(): TagCollectionObj { return (this.tagsObj ??= new TagCollectionObj(this.context, () => this.data.tags)) }
  get placeholderFormat(): PlaceholderFormatObj {
    this.context.gate('1.8', 'Shape.placeholderFormat')
    return (this.placeholderObj ??= new PlaceholderFormatObj(this.context, () => this.data))
  }
  get group(): ShapeGroupObj {
    this.context.gate('1.8', 'Shape.group')
    return (this.groupObj ??= new ShapeGroupObj(this.context, this.model, () => this.data))
  }
  get parentGroup(): ShapeObj {
    this.context.gate('1.8', 'Shape.parentGroup')
    const parent = this.shape ? locate(this.model, this.shape)?.parent ?? null : null
    return new ShapeObj(this.context, this.model, parent)
  }

  protected override onLoaded(head: string, rest: string): void {
    const child: Record<string, (() => FakeObject) | undefined> = {
      fill: () => this.fill, lineFormat: () => this.lineFormat, textFrame: () => this.textFrame, tags: () => this.tags,
      placeholderFormat: () => this.placeholderFormat, group: () => this.group,
    }
    child[head]?.().markLoaded(rest || '*')
  }

  delete(): void {
    this.write(() => {
      const { list } = this.where()
      list.splice(list.indexOf(this.data), 1)
      const sel = this.model.selection
      this.model.selection = { ...sel, shapeIds: sel.shapeIds.filter((id) => id !== this.data.id) }
    })
  }

  setZOrder(position: string): void {
    this.context.gate('1.8', 'Shape.setZOrder')
    this.write(() => {
      const { list } = this.where()
      const i = list.indexOf(this.data)
      const to = { BringToFront: list.length - 1, SendToBack: 0, BringForward: Math.min(i + 1, list.length - 1), SendBackward: Math.max(i - 1, 0) }[position]
      if (to === undefined) throw new FakeOfficeError('InvalidArgument', `Unknown z-order position ${position}`)
      list.splice(i, 1)
      list.splice(to, 0, this.data)
    })
  }

  getTable(): TableObj {
    this.context.gate('1.8', 'Shape.getTable')
    return new TableObj(this.context, this.model, () => this.data)
  }
  getParentSlide(): SlideObj {
    this.context.gate('1.5', 'Shape.getParentSlide')
    return new SlideObj(this.context, this.model, () => this.where().slide)
  }
  getTextFrameOrNullObject(): TextFrameObj {
    this.context.gate('1.10', 'Shape.getTextFrameOrNullObject')
    return new TextFrameObj(this.context, { model: this.model, shape: () => this.data }, true)
  }
}

type AddOptions = { left?: number; top?: number; width?: number; height?: number }

const position = (o: AddOptions | undefined): AddOptions => {
  const out: AddOptions = {}
  for (const k of ['left', 'top', 'width', 'height'] as const) if (typeof o?.[k] === 'number') out[k] = o[k]
  return out
}

abstract class ShapeListObj extends FakeCollection<ShapeObj> {
  constructor(context: Ctx, readonly model: FakePresentation) {
    super(context)
  }
  protected abstract list(): FakeShape[]
  protected source(): object[] { return this.list() }
  protected wrap(data: object): ShapeObj { return new ShapeObj(this.context, this.model, data as FakeShape) }
  protected override itemVersions(): Record<string, string> { return SHAPE_VERSIONS }
  private find(id: string): FakeShape | null { return this.list().find((s) => s.id === id) ?? null }
  getItem(id: string): ShapeObj { return new ShapeObj(this.context, this.model, this.find(id)) }
  getItemOrNullObject(id: string): ShapeObj { return new ShapeObj(this.context, this.model, this.find(id), true) }
  getItemAt(index: number): ShapeObj { return new ShapeObj(this.context, this.model, this.list()[index] ?? null) }
  getCount(): FakeClientResult<number> { return new FakeClientResult(this.context, () => this.list().length) }
}

/** A slide's shapes: lists the top-level shapes and adds new ones in front. */
export class ShapeCollectionObj extends ShapeListObj {
  constructor(context: Ctx, model: FakePresentation, private readonly slide: () => FakeSlide) {
    super(context, model)
  }
  protected list(): FakeShape[] { return this.slide().shapes }
  private add(data: FakeShape, version: string, what: string, check?: () => void): ShapeObj {
    this.context.gate(version, what)
    this.write(() => {
      check?.()
      this.slide().shapes.push(data)
    })
    return this.proxy(data)
  }
  addGeometricShape(type: string, options?: AddOptions): ShapeObj {
    return this.add(shapeData({ geometricShapeType: type, ...position(options) }), '1.4', 'ShapeCollection.addGeometricShape')
  }
  addTextBox(text: string, options?: AddOptions): ShapeObj {
    return this.add(textBoxData(text, position(options)), '1.4', 'ShapeCollection.addTextBox')
  }
  addLine(connectorType = 'Straight', options?: AddOptions): ShapeObj {
    return this.add(lineData({ connectorType, ...position(options) }), '1.4', 'ShapeCollection.addLine')
  }
  addTable(rows: number, columns: number, options: Record<string, unknown> & AddOptions = {}): ShapeObj {
    if (options.style !== undefined) this.context.gate('1.9', 'TableAddOptions.style')
    const values = (options.values as string[][] | undefined) ?? Array.from({ length: rows }, () => Array.from({ length: columns }, () => ''))
    const data = tableShape(values.map((r) => [...r]), position(options))
    if (data.table) data.table.options = options
    const fits = (grid: unknown): boolean => !Array.isArray(grid) || (grid.length === rows && grid.every((r) => Array.isArray(r) && r.length === columns))
    return this.add(data, '1.8', 'ShapeCollection.addTable', () => {
      if (rows < 1 || columns < 1 || !fits(options.values) || !fits(options.specificCellProperties)) {
        throw new FakeOfficeError('InvalidArgument', `Table values or cell properties do not match ${rows} by ${columns}`)
      }
    })
  }
  addGroup(values: Array<string | ShapeObj>): ShapeObj {
    this.context.gate('1.8', 'ShapeCollection.addGroup')
    const g = groupData([])
    this.write(() => {
      const list = this.slide().shapes
      const members = values.map((v) => (typeof v === 'string' ? list.find((s) => s.id === v) : v.data))
      if (members.some((m) => !m)) throw new FakeOfficeError('InvalidArgument', 'addGroup: a shape is not on this slide')
      makeGroup(this.model, list, members as FakeShape[], g)
    })
    return this.proxy(g)
  }
}

/** The selection, or a group's members. */
export class ShapeScopedCollectionObj extends ShapeListObj {
  constructor(context: Ctx, model: FakePresentation, private readonly shapes: () => FakeShape[]) {
    super(context, model)
  }
  protected list(): FakeShape[] { return this.shapes() }
  group(): ShapeObj {
    this.context.gate('1.8', 'ShapeScopedCollection.group')
    const g = groupData([])
    this.write(() => {
      const members = this.shapes()
      const first = members[0]
      const where = first ? locate(this.model, first) : null
      if (!where) throw new FakeOfficeError('InvalidArgument', 'Nothing to group')
      makeGroup(this.model, where.list, members, g)
    })
    return new ShapeObj(this.context, this.model, g)
  }
}
