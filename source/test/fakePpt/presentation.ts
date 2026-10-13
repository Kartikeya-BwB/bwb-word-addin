// Fake presentation, slides, layouts, masters and page setup.

import { Ctx, FakeClientResult, FakeCollection, FakeObject, FakeOfficeError } from './base.ts'
import type { FakeLayout, FakeMaster, FakePresentation, FakeShape, FakeSlide } from './model.ts'
import { nextId, slide as slideData } from './model.ts'
import { allShapes, findShape, findSlide, selectedShapes } from './queries.ts'
import { ShapeCollectionObj, ShapeScopedCollectionObj, TagCollectionObj } from './shapes.ts'
import { TextRangeObj } from './text.ts'

/** A fresh copy of a layout placeholder, with its own id and text, for a slide added with that layout. */
function copyShape(s: FakeShape): FakeShape {
  return {
    ...s,
    id: nextId(),
    fill: { ...s.fill },
    line: { ...s.line },
    tags: new Map(s.tags),
    textFrame: s.textFrame ? { ...s.textFrame, runs: s.textFrame.runs.map((r) => ({ text: r.text, font: { ...r.font } })) } : null,
    ...(s.members ? { members: s.members.map(copyShape) } : {}),
  }
}

class SlideLayoutObj extends FakeObject {
  constructor(context: Ctx, private readonly layout: () => FakeLayout | undefined) {
    super(context)
  }
  private data(): FakeLayout {
    const l = this.layout()
    if (!l) throw new FakeOfficeError('ItemNotFound', 'The layout was not found')
    return l
  }
  protected override checkLoad(): void { this.data() }
  protected override versionFor(path: string): string | undefined { return path === 'type' ? '1.8' : undefined }
  get id(): string { this.need('id'); return this.data().id }
  get name(): string { this.need('name'); return this.data().name }
  get type(): string { this.need('type'); return this.data().type }
}

class SlideLayoutCollectionObj extends FakeCollection<SlideLayoutObj> {
  constructor(context: Ctx, private readonly master: () => FakeMaster) {
    super(context)
  }
  protected source(): object[] { return this.master().layouts }
  protected wrap(data: object): SlideLayoutObj { return new SlideLayoutObj(this.context, () => data as FakeLayout) }
  getItem(id: string): SlideLayoutObj { return new SlideLayoutObj(this.context, () => this.master().layouts.find((l) => l.id === id)) }
  getItemAt(index: number): SlideLayoutObj { return new SlideLayoutObj(this.context, () => this.master().layouts[index]) }
  getCount(): FakeClientResult<number> { return new FakeClientResult(this.context, () => this.master().layouts.length) }
}

class SlideMasterObj extends FakeObject {
  constructor(context: Ctx, private readonly master: () => FakeMaster | undefined) {
    super(context)
  }
  private data(): FakeMaster {
    const m = this.master()
    if (!m) throw new FakeOfficeError('ItemNotFound', 'The slide master was not found')
    return m
  }
  protected override checkLoad(): void { this.data() }
  get id(): string { this.need('id'); return this.data().id }
  get name(): string { this.need('name'); return this.data().name }
  get layouts(): SlideLayoutCollectionObj { return new SlideLayoutCollectionObj(this.context, () => this.data()) }
  protected override onLoaded(head: string, rest: string): void {
    if (head === 'layouts') this.layouts.markLoaded(rest || '*')
  }
}

class SlideMasterCollectionObj extends FakeCollection<SlideMasterObj> {
  constructor(context: Ctx, private readonly model: FakePresentation) {
    super(context)
  }
  protected source(): object[] { return this.model.masters }
  protected wrap(data: object): SlideMasterObj { return new SlideMasterObj(this.context, () => data as FakeMaster) }
  getItem(id: string): SlideMasterObj { return new SlideMasterObj(this.context, () => this.model.masters.find((m) => m.id === id)) }
  getItemAt(index: number): SlideMasterObj { return new SlideMasterObj(this.context, () => this.model.masters[index]) }
}

export class SlideObj extends FakeObject {
  private shapesObj: ShapeCollectionObj | undefined
  private tagsObj: TagCollectionObj | undefined

  constructor(context: Ctx, private readonly model: FakePresentation, private readonly slide: () => FakeSlide | undefined, private readonly nullable = false) {
    super(context)
  }
  /** The slide behind this proxy; fails the sync (ItemNotFound) when there is none. */
  data(): FakeSlide {
    const s = this.slide()
    if (!s) throw new FakeOfficeError('ItemNotFound', 'The slide was not found')
    return s
  }
  protected override isNull(): boolean { return this.nullable && !this.slide() }
  protected override checkLoad(): void { if (!this.nullable) this.data() }
  protected override versionFor(path: string): string | undefined { return path === 'index' ? '1.8' : undefined }

  get id(): string { this.need('id'); return this.data().id }
  get index(): number { this.need('index'); return this.model.slides.indexOf(this.data()) }
  get shapes(): ShapeCollectionObj { return (this.shapesObj ??= new ShapeCollectionObj(this.context, this.model, () => this.data())) }
  get tags(): TagCollectionObj { return (this.tagsObj ??= new TagCollectionObj(this.context, () => this.data().tags)) }
  get layout(): SlideLayoutObj {
    return new SlideLayoutObj(this.context, () => this.model.masters.flatMap((m) => m.layouts).find((l) => l.id === this.data().layoutId))
  }
  get slideMaster(): SlideMasterObj { return new SlideMasterObj(this.context, () => this.model.masters.find((m) => m.id === this.data().masterId)) }

  protected override onLoaded(head: string, rest: string): void {
    if (head === 'shapes') this.shapes.markLoaded(rest || '*')
    if (head === 'tags') this.tags.markLoaded(rest || '*')
  }

  delete(): void {
    this.write(() => {
      const s = this.data()
      this.model.slides.splice(this.model.slides.indexOf(s), 1)
      const sel = this.model.selection
      if (sel.slideIds.includes(s.id)) {
        const rest = sel.slideIds.filter((id) => id !== s.id)
        const fallback = this.model.slides[0]?.id
        this.model.selection = { slideIds: rest.length ? rest : fallback ? [fallback] : [], shapeIds: [], textRange: null }
      }
    })
  }
  moveTo(index: number): void {
    this.context.gate('1.8', 'Slide.moveTo')
    this.write(() => {
      const s = this.data()
      if (index < 0 || index >= this.model.slides.length) throw new FakeOfficeError('InvalidArgument', `No slide position ${index}`)
      this.model.slides.splice(this.model.slides.indexOf(s), 1)
      this.model.slides.splice(index, 0, s)
    })
  }
  setSelectedShapes(shapeIds: string[]): void {
    this.context.gate('1.5', 'Slide.setSelectedShapes')
    this.write(() => {
      const s = this.data()
      const missing = shapeIds.find((id) => !allShapes(s).some((x) => x.id === id))
      if (missing) throw new FakeOfficeError('InvalidArgument', `Shape ${missing} is not on slide ${s.id}`)
      this.model.selection = { slideIds: [s.id], shapeIds: [...shapeIds], textRange: null }
    })
  }
  applyLayout(layout: SlideLayoutObj): void {
    this.context.gate('1.8', 'Slide.applyLayout')
    this.write(() => {
      this.data().layoutId = layout.id
    })
  }
  exportAsBase64(): FakeClientResult<string> {
    this.context.gate('1.8', 'Slide.exportAsBase64')
    return new FakeClientResult(this.context, () => btoa(`fake-slide:${this.data().id}`))
  }
}

/** Records an export of these slides and returns a stand-in for the .pptx, as base64. */
function exportSlides(model: FakePresentation, ids: string[]): string {
  const missing = ids.find((id) => !findSlide(model, id))
  if (missing) throw new FakeOfficeError('ItemNotFound', `No slide ${missing}`)
  model.slideExports.push(ids)
  return btoa(`fake-presentation:${ids.join(',')}`)
}

abstract class SlideListObj extends FakeCollection<SlideObj> {
  constructor(context: Ctx, protected readonly model: FakePresentation) {
    super(context)
  }
  protected abstract list(): FakeSlide[]
  protected source(): object[] { return this.list() }
  protected wrap(data: object): SlideObj { return new SlideObj(this.context, this.model, () => data as FakeSlide) }
  protected override itemVersions(): Record<string, string> { return { index: '1.8' } }
  getItem(id: string): SlideObj { return new SlideObj(this.context, this.model, () => this.list().find((s) => s.id === id)) }
  getItemOrNullObject(id: string): SlideObj { return new SlideObj(this.context, this.model, () => this.list().find((s) => s.id === id), true) }
  getItemAt(index: number): SlideObj { return new SlideObj(this.context, this.model, () => this.list()[index]) }
  getCount(): FakeClientResult<number> { return new FakeClientResult(this.context, () => this.list().length) }
}

/** presentation.slides: every slide in order; add() appends a slide made from a layout. */
class SlideCollectionObj extends SlideListObj {
  protected list(): FakeSlide[] { return this.model.slides }
  exportAsBase64Presentation(values: Array<string | SlideObj>): FakeClientResult<string> {
    this.context.gate('1.10', 'SlideCollection.exportAsBase64Presentation')
    return new FakeClientResult(this.context, () => exportSlides(this.model, values.map((v) => (typeof v === 'string' ? v : v.data().id))))
  }
  add(options: { layoutId?: string; slideMasterId?: string } = {}): void {
    this.context.gate('1.3', 'SlideCollection.add')
    this.write(() => {
      const master = this.model.masters.find((m) => m.id === options.slideMasterId) ?? this.model.masters[0]
      const layout = options.layoutId ? this.model.masters.flatMap((m) => m.layouts).find((l) => l.id === options.layoutId) : master?.layouts[0]
      if (!master || !layout) throw new FakeOfficeError('InvalidArgument', `Unknown layout ${options.layoutId ?? ''}`)
      const owner = this.model.masters.find((m) => m.layouts.includes(layout)) ?? master
      this.model.slides.push(slideData(layout.placeholders.map(copyShape), { layoutId: layout.id, masterId: owner.id }))
    })
  }
}

/** getSelectedSlides(): the selected slides, the active slide first. */
class SlideScopedCollectionObj extends SlideListObj {
  protected list(): FakeSlide[] {
    return this.model.selection.slideIds.map((id) => findSlide(this.model, id)).filter((s): s is FakeSlide => s !== undefined)
  }
  exportAsBase64Presentation(): FakeClientResult<string> {
    this.context.gate('1.10', 'SlideScopedCollection.exportAsBase64Presentation')
    return new FakeClientResult(this.context, () => exportSlides(this.model, this.list().map((s) => s.id)))
  }
}

class PageSetupObj extends FakeObject {
  constructor(context: Ctx, private readonly model: FakePresentation) {
    super(context)
  }
  get slideWidth(): number { this.need('slideWidth'); return this.model.pageSetup.slideWidth }
  set slideWidth(v: number) { this.write(() => (this.model.pageSetup.slideWidth = v)) }
  get slideHeight(): number { this.need('slideHeight'); return this.model.pageSetup.slideHeight }
  set slideHeight(v: number) { this.write(() => (this.model.pageSetup.slideHeight = v)) }
}

export class PresentationObj extends FakeObject {
  private tagsObj: TagCollectionObj | undefined
  private slidesObj: SlideCollectionObj | undefined

  constructor(context: Ctx, private readonly model: FakePresentation) {
    super(context)
  }
  protected override versionFor(path: string): string | undefined { return path.startsWith('pageSetup') ? '1.10' : path === 'id' ? '1.5' : undefined }

  get id(): string { this.need('id'); return this.model.id }
  get title(): string { this.need('title'); return this.model.title }
  get slides(): SlideCollectionObj { return (this.slidesObj ??= new SlideCollectionObj(this.context, this.model)) }
  get tags(): TagCollectionObj { return (this.tagsObj ??= new TagCollectionObj(this.context, () => this.model.tags)) }
  get slideMasters(): SlideMasterCollectionObj { return new SlideMasterCollectionObj(this.context, this.model) }
  get pageSetup(): PageSetupObj {
    this.context.gate('1.10', 'Presentation.pageSetup')
    return new PageSetupObj(this.context, this.model)
  }

  protected override onLoaded(head: string, rest: string): void {
    if (head === 'slides') this.slides.markLoaded(rest || '*')
    if (head === 'tags') this.tags.markLoaded(rest || '*')
  }

  getSelectedShapes(): ShapeScopedCollectionObj {
    this.context.gate('1.5', 'Presentation.getSelectedShapes')
    return new ShapeScopedCollectionObj(this.context, this.model, () => selectedShapes(this.model))
  }
  getSelectedSlides(): SlideScopedCollectionObj {
    this.context.gate('1.5', 'Presentation.getSelectedSlides')
    return new SlideScopedCollectionObj(this.context, this.model)
  }
  private selectedText(nullable: boolean): TextRangeObj | NullTextRangeObj {
    this.context.gate('1.5', 'Presentation.getSelectedTextRange')
    const t = this.model.selection.textRange
    const shape = t ? findShape(this.model, t.shapeId) : undefined
    if (!t || !shape) return new NullTextRangeObj(this.context, nullable)
    return new TextRangeObj(this.context, { model: this.model, shape: () => shape }, t.start, t.length)
  }
  getSelectedTextRange(): TextRangeObj | NullTextRangeObj { return this.selectedText(false) }
  getSelectedTextRangeOrNullObject(): TextRangeObj | NullTextRangeObj { return this.selectedText(true) }
  setSelectedSlides(slideIds: string[]): void {
    this.context.gate('1.5', 'Presentation.setSelectedSlides')
    this.write(() => {
      const missing = slideIds.find((id) => !findSlide(this.model, id))
      if (missing) throw new FakeOfficeError('InvalidArgument', `No slide ${missing}`)
      this.model.selection = { slideIds: [...slideIds], shapeIds: [], textRange: null }
    })
  }
  insertSlidesFromBase64(base64: string, options: Record<string, unknown> = {}): void {
    this.context.gate('1.2', 'Presentation.insertSlidesFromBase64')
    this.write(() => {
      const target = typeof options.targetSlideId === 'string' ? options.targetSlideId : undefined
      if (target !== undefined && !findSlide(this.model, target)) throw new FakeOfficeError('SlideNotFound', `No slide ${target}`)
      this.model.slideInserts.push({ base64, options })
      const made = this.model.slidesFromBase64?.(base64, options) ?? []
      const at = target === undefined ? 0 : this.model.slides.findIndex((s) => s.id === target) + 1
      this.model.slides.splice(at, 0, ...made)
    })
  }
}

/** What getSelectedTextRange(OrNullObject) returns when there is no text selection. */
class NullTextRangeObj extends FakeObject {
  constructor(context: Ctx, private readonly nullable: boolean) {
    super(context)
  }
  protected override isNull(): boolean { return true }
  protected override checkLoad(): void {
    if (!this.nullable) throw new FakeOfficeError('InvalidSelection', 'There is no text selection')
  }
}
