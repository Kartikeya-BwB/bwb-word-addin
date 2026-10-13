// The fake Word and Office globals: install one over a FakeModel, run the add-in's Word code, inspect the model.

import { Ctx, FakeCollection, FakeObject } from './base.ts'
import type { FakeModel, FakePart } from './model.ts'
import type { Container } from './objects.ts'
import {
  ControlCollection, ParagraphCollection, ParagraphObj, RangeCollection, RangeObj, TableCollection, insertOoxmlAt, partContainer,
} from './objects.ts'
import { para } from './model.ts'
import { PictureCollection } from './pictures.ts'
import { BookmarkRangeObj } from './ranges.ts'
import { CommentCollection, FieldCollection, TrackedChangeCollection, documentReview } from './review.ts'
import { SelectionObj } from './selection.ts'

class BodyObj extends FakeObject {
  constructor(context: Ctx, private readonly container: Container, private readonly part?: FakePart) {
    super(context)
  }
  get text(): string { this.need('text'); return this.container.list.map((p) => p.text).join('\r') }
  get paragraphs(): ParagraphCollection { return new ParagraphCollection(this.context, this.container, () => this.container.list) }
  get tables(): TableCollection { return new TableCollection(this.context, this.container) }
  get contentControls(): ControlCollection { return new ControlCollection(this.context, this.container) }
  get inlinePictures(): PictureCollection {
    const part = this.part
    const model = this.container.model
    return new PictureCollection(this.context, this.container, () =>
      part ? Array.from({ length: part.pictures }, () => ({ altTextTitle: 'logo', altTextDescription: '' })) : model.pictures,
    )
  }
  get fields(): FieldCollection { return new FieldCollection(this.context, this.container.model) }
  search(needle: string, options: { matchCase?: boolean } = {}): RangeCollection {
    return new RangeCollection(this.context, () =>
      this.container.list.flatMap((p) => {
        const hay = options.matchCase ? p.text : p.text.toLowerCase()
        const n = options.matchCase ? needle : needle.toLowerCase()
        const out: RangeObj[] = []
        for (let i = hay.indexOf(n); n && i !== -1; i = hay.indexOf(n, i + 1)) out.push(new RangeObj(this.context, this.container, p, i, i + n.length))
        return out
      }),
    )
  }
  getComments(): CommentCollection { return new CommentCollection(this.context, this.container.model) }
  getTrackedChanges(): TrackedChangeCollection { return new TrackedChangeCollection(this.context, this.container.model) }
  getRange(location = 'Whole'): RangeObj {
    const list = this.container.list
    const p = location === 'End' ? list[list.length - 1] : list[0]
    const at = location === 'End' ? (p?.text.length ?? 0) : 0
    return new RangeObj(this.context, this.container, p ?? null, at, location === 'Whole' ? (p?.text.length ?? 0) : at)
  }
  insertOoxml(ooxml: string, location: string): RangeObj {
    return insertOoxmlAt(this.context, this.container, null, ooxml, location)
  }
  insertParagraph(text: string, location: string): ParagraphObj {
    const p = para(text)
    this.context.queue(() => (location === 'Start' ? this.container.list.unshift(p) : this.container.list.push(p)))
    return new ParagraphObj(this.context, this.container, p)
  }
  getOoxml(): { value: string } {
    const result = { value: '' }
    this.context.queue(() => (result.value = toFlatOpc(this.container)))
    return result
  }
}

function toFlatOpc(container: Container): string {
  const esc = (s: string): string => s.replace(/&/g, '&amp;').replace(/</g, '&lt;')
  const body = container.list
    .map((p) => {
      const f = p.font
      const rPr =
        (f.name ? `<w:rFonts w:ascii="${f.name}"/>` : '') +
        (f.color ? `<w:color w:val="${f.color.replace('#', '')}"/>` : '') +
        (f.highlightColor ? `<w:highlight w:val="${f.highlightColor}"/>` : '') +
        (f.size ? `<w:sz w:val="${f.size * 2}"/>` : '')
      return `<w:p><w:r><w:rPr>${rPr}</w:rPr><w:t>${esc(p.text)}</w:t></w:r></w:p>`
    })
    .join('')
  return (
    '<pkg:package xmlns:pkg="http://schemas.microsoft.com/office/2006/xmlPackage">' +
    '<pkg:part pkg:name="/word/document.xml"><pkg:xmlData>' +
    `<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:body>${body}</w:body></w:document>` +
    '</pkg:xmlData></pkg:part></pkg:package>'
  )
}

class StyleObj extends FakeObject {
  constructor(context: Ctx, private readonly model: FakeModel, private readonly name: string) {
    super(context)
  }
  get isNullObject(): boolean { this.need('isNullObject'); return !this.model.styles.has(this.name) }
  get font(): { set: (v: unknown) => void } {
    return { set: (v: unknown) => this.context.queue(() => this.model.styleUpdates.push({ name: this.name, font: v })) }
  }
  get paragraphFormat(): { set: (v: unknown) => void } {
    return { set: (v: unknown) => this.context.queue(() => this.model.styleUpdates.push({ name: this.name, paragraph: v })) }
  }
}

class CustomProp extends FakeObject {
  constructor(context: Ctx, private readonly model: FakeModel, readonly keyName: string) {
    super(context)
  }
  get isNullObject(): boolean { this.need('isNullObject'); return !this.model.custom.has(this.keyName) }
  get key(): string { this.need('key'); return this.keyName }
  get value(): string { this.need('value'); return this.model.custom.get(this.keyName) ?? '' }
  delete(): void { this.context.queue(() => this.model.custom.delete(this.keyName)) }
}

class CustomProps extends FakeCollection<CustomProp> {
  constructor(context: Ctx, private readonly model: FakeModel) {
    super(context)
  }
  protected make(): CustomProp[] {
    return [...this.model.custom.keys()].map((k) => new CustomProp(this.context, this.model, k))
  }
  add(key: string, value: string): void { this.context.queue(() => this.model.custom.set(key, value)) }
  getItemOrNullObject(key: string): CustomProp { return new CustomProp(this.context, this.model, key) }
}

class PropertiesObj extends FakeObject {
  constructor(context: Ctx, private readonly model: FakeModel) {
    super(context)
  }
  private get p(): FakeModel['properties'] { return this.model.properties }
  get title(): string { this.need('title'); return this.p.title }
  set title(v: string) { this.context.queue(() => (this.p.title = v)) }
  get subject(): string { this.need('subject'); return this.p.subject }
  set subject(v: string) { this.context.queue(() => (this.p.subject = v)) }
  get author(): string { this.need('author'); return this.p.author }
  set author(v: string) { this.context.queue(() => (this.p.author = v)) }
  get company(): string { this.need('company'); return this.p.company }
  set company(v: string) { this.context.queue(() => (this.p.company = v)) }
  get customProperties(): CustomProps { return new CustomProps(this.context, this.model) }
}

class SectionObj extends FakeObject {
  constructor(context: Ctx, private readonly model: FakeModel, private readonly index: number) {
    super(context)
  }
  private part(kind: 'headers' | 'footers', type: string): BodyObj {
    const section = this.model.sections[this.index]
    const part = section?.[kind][type as 'Primary']
    if (!part) throw new Error(`no ${kind} ${type}`)
    return new BodyObj(this.context, partContainer(this.model, part), part)
  }
  getHeader(type: string): BodyObj { return this.part('headers', type) }
  getFooter(type: string): BodyObj { return this.part('footers', type) }
}

class SectionCollection extends FakeCollection<SectionObj> {
  constructor(context: Ctx, private readonly model: FakeModel) {
    super(context)
  }
  protected make(): SectionObj[] {
    return this.model.sections.map((_, i) => new SectionObj(this.context, this.model, i))
  }
}

class DocumentObj extends FakeObject {
  readonly body: BodyObj
  readonly sections: SectionCollection
  readonly properties: PropertiesObj
  private readonly selection: SelectionObj
  constructor(context: Ctx, private readonly model: FakeModel) {
    super(context)
    const container: Container = { model, list: model.paragraphs, controls: model.controls }
    this.body = new BodyObj(context, container)
    this.sections = new SectionCollection(context, model)
    this.properties = new PropertiesObj(context, model)
    this.selection = new SelectionObj(context, container, model)
    Object.assign(this, documentReview(context, model))
  }
  getSelection(): SelectionObj { return this.selection }
  getStyles(): { getByNameOrNullObject: (name: string) => StyleObj } {
    return { getByNameOrNullObject: (name: string) => new StyleObj(this.context, this.model, name) }
  }
  get changeTrackingMode(): string { this.need('changeTrackingMode'); return this.model.changeTrackingMode ?? 'Off' }
  set changeTrackingMode(v: string) { this.context.queue(() => (this.model.changeTrackingMode = v)) }
  getBookmarkRangeOrNullObject(name: string): BookmarkRangeObj { return new BookmarkRangeObj(this.context, this.model, name) }
}

export function installFakeWord(model: FakeModel): { runs: () => number } {
  let runs = 0
  const run = async <T>(fn: (context: unknown) => Promise<T>): Promise<T> => {
    runs += 1
    const ctx = new Ctx()
    const context = {
      sync: () => ctx.sync(),
      document: new DocumentObj(ctx, model),
      application: { createDocument: (b64: string) => ({ open: () => ctx.queue(() => model.created.push(b64)) }) },
    }
    const result = await fn(context)
    await ctx.sync()
    return result
  }
  const g = globalThis as Record<string, unknown>
  g.Word = { run }
  g.Office = {
    AsyncResultStatus: { Succeeded: 'succeeded', Failed: 'failed' },
    EventType: { DocumentSelectionChanged: 'documentSelectionChanged' },
    PlatformType: { PC: 'PC', OfficeOnline: 'OfficeOnline', Mac: 'Mac', iOS: 'iOS', Android: 'Android', Universal: 'Universal' },
    context: {
      get platform() { return model.platform },
      requirements: {
        isSetSupported: (set: string, version: string) =>
          Number(version) <= (set === 'WordApiDesktop' && model.desktopApiVersion !== undefined ? model.desktopApiVersion : model.apiVersion),
      },
      document: {
        get url() { return model.url },
        settings: {
          get: (k: string) => model.settings.get(k),
          set: (k: string, v: unknown) => model.settings.set(k, v),
          saveAsync: (cb: (r: { status: string }) => void) => cb({ status: 'succeeded' }),
        },
        addHandlerAsync: (type: string, handler: () => void, cb?: (r: { status: string }) => void) => {
          model.handlers = [...(model.handlers ?? []), { type, handler }]
          cb?.({ status: 'succeeded' })
        },
        removeHandlerAsync: (type: string, options?: { handler?: () => void }, cb?: (r: { status: string }) => void) => {
          model.handlers = (model.handlers ?? []).filter((h) => h.type !== type || (options?.handler && h.handler !== options.handler))
          cb?.({ status: 'succeeded' })
        },
      },
    },
  }
  return { runs: () => runs }
}

export function uninstallFakeWord(): void {
  const g = globalThis as Record<string, unknown>
  delete g.Word
  delete g.Office
}
