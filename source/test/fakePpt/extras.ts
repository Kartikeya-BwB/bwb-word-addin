// Fake slide, layout and master backgrounds (PowerPointApi 1.10) and slide hyperlinks (1.6), added for the slide tools
// and the Pre-QA check. A background missing from the model follows the master's; a master's is solid white.

import { Ctx, FakeClientResult, FakeCollection, FakeObject, FakeOfficeError } from './base.ts'
import type { FakeBackground, FakeHyperlink, FakePresentation, FakeSlide } from './model.ts'
import { allShapes } from './queries.ts'
import { ShapeObj } from './shapes.ts'
import { normaliseColour } from './text.ts'

const WHITE: FakeBackground = { type: 'Solid', color: '#FFFFFF', followsMaster: false }

/** The background a slide shows: its own, else its layout's, else its master's. */
export function effectiveBackground(model: FakePresentation, slide: FakeSlide): FakeBackground {
  if (slide.background && !slide.background.followsMaster) return slide.background
  const master = model.masters.find((m) => m.id === slide.masterId)
  const layout = model.masters.flatMap((m) => m.layouts).find((l) => l.id === slide.layoutId)
  if (layout?.background && !layout.background.followsMaster) return layout.background
  return master?.background ?? WHITE
}

class SolidFillObj extends FakeObject {
  constructor(context: Ctx, private readonly fill: () => FakeBackground | undefined) {
    super(context)
  }
  protected override isNull(): boolean { return !this.fill() }
  get color(): string { this.need('color'); return normaliseColour(this.fill()?.color ?? '') }
  get transparency(): number { this.need('transparency'); return 0 }
}

class BackgroundFillObj extends FakeObject {
  constructor(context: Ctx, private readonly fill: () => FakeBackground) {
    super(context)
  }
  get type(): string { this.need('type'); return this.fill().type }
  getSolidFillOrNullObject(): SolidFillObj {
    return new SolidFillObj(this.context, () => (this.fill().type === 'Solid' ? this.fill() : undefined))
  }
}

/** Slide.background, SlideLayout.background or SlideMaster.background. */
export class BackgroundObj extends FakeObject {
  constructor(context: Ctx, private readonly own: () => FakeBackground | undefined, private readonly effective: () => FakeBackground) {
    super(context)
  }
  get fill(): BackgroundFillObj { return new BackgroundFillObj(this.context, this.effective) }
  get isMasterBackgroundFollowed(): boolean { this.need('isMasterBackgroundFollowed'); return this.own()?.followsMaster ?? true }
  get areBackgroundGraphicsHidden(): boolean { this.need('areBackgroundGraphicsHidden'); return false }
  protected override onLoaded(head: string, rest: string): void {
    if (head === 'fill') this.fill.markLoaded(rest || '*')
  }
}

class HyperlinkObj extends FakeObject {
  constructor(context: Ctx, private readonly model: FakePresentation, private readonly slide: () => FakeSlide, private readonly link: FakeHyperlink) {
    super(context)
  }
  protected override versionFor(path: string): string | undefined { return path === 'type' ? '1.10' : undefined }
  get address(): string { this.need('address'); return this.link.address }
  get screenTip(): string { this.need('screenTip'); return this.link.screenTip ?? '' }
  get type(): string { this.need('type'); return this.link.shapeId ? 'Shape' : 'TextRange' }
  getLinkedShapeOrNullObject(): ShapeObj {
    this.context.gate('1.10', 'Hyperlink.getLinkedShapeOrNullObject')
    const id = this.link.shapeId
    const shape = id === undefined ? null : allShapes(this.slide()).find((s) => s.id === id) ?? null
    return new ShapeObj(this.context, this.model, shape, true)
  }
}

/** Slide.hyperlinks: the links on one slide, from FakeSlide.hyperlinks. */
export class HyperlinkCollectionObj extends FakeCollection<HyperlinkObj> {
  constructor(context: Ctx, private readonly model: FakePresentation, private readonly slide: () => FakeSlide) {
    super(context)
  }
  protected source(): object[] { return this.slide().hyperlinks ?? [] }
  protected wrap(data: object): HyperlinkObj { return new HyperlinkObj(this.context, this.model, this.slide, data as FakeHyperlink) }
  protected override itemVersions(): Record<string, string> { return { type: '1.10' } }
  getCount(): FakeClientResult<number> { return new FakeClientResult(this.context, () => this.source().length) }
  getItemAt(index: number): HyperlinkObj {
    const link = this.slide().hyperlinks?.[index]
    if (!link) throw new FakeOfficeError('InvalidArgument', `No hyperlink ${index}`)
    return new HyperlinkObj(this.context, this.model, this.slide, link)
  }
}
