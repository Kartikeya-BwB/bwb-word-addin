// The fake Word and Office globals: install one over a FakeModel, run the add-in's Word code, inspect the model.

import { Ctx, FakeCollection, FakeObject } from './base.ts'
import type { FakeModel, FakePart } from './model.ts'
import type { Container } from './objects.ts'
import {
  ControlCollection, NullTableObj, ParagraphCollection, ParagraphObj, RangeCollection, RangeObj, TableCollection, TableObj,
  insertOoxmlAt, partContainer,
} from './objects.ts'
import { para } from './model.ts'

class Simple extends FakeObject {}

class CountCollection extends FakeCollection<Simple> {
  constructor(context: Ctx, private readonly count: () => number) {
    super(context)
  }
  protected make(): Simple[] {
    return Array.from({ length: this.count() }, () => new Simple(this.context))
  }
}

class FieldObj extends FakeObject {
  constructor(context: Ctx, private readonly data: { code: string; updates: number }) {
    super(context)
  }
  get code(): string { this.need('code'); return this.data.code }
  updateResult(): void { this.context.queue(() => (this.data.updates += 1)) }
}

class FieldCollection extends FakeCollection<FieldObj> {
  constructor(context: Ctx, private readonly model: FakeModel) {
    super(context)
  }
  protected make(): FieldObj[] {
    return this.model.fields.map((f) => new FieldObj(this.context, f))
  }
}

interface PictureData {
  altTextTitle: string
  altTextDescription: string
}

class PictureObj extends FakeObject {
  constructor(context: Ctx, private readonly model: FakeModel, private readonly data: PictureData) {
    super(context)
  }
  get altTextTitle(): string { this.need('altTextTitle'); return this.data.altTextTitle }
  get altTextDescription(): string { this.need('altTextDescription'); return this.data.altTextDescription }
  select(): void { this.context.queue(() => this.model.selected.push(`picture:${this.data.altTextTitle}`)) }
}

class PictureCollection extends FakeCollection<PictureObj> {
  constructor(context: Ctx, private readonly model: FakeModel, private readonly list: () => PictureData[]) {
    super(context)
  }
  protected make(): PictureObj[] {
    return this.list().map((p) => new PictureObj(this.context, this.model, p))
  }
}

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
    return new PictureCollection(this.context, model, () =>
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
  getComments(): CountCollection { return new CountCollection(this.context, () => this.container.model.comments) }
  getTrackedChanges(): CountCollection { return new CountCollection(this.context, () => this.container.model.trackedChanges) }
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

export function installFakeWord(model: FakeModel): { runs: () => number } {
  let runs = 0
  const body = (ctx: Ctx): BodyObj => new BodyObj(ctx, { model, list: model.paragraphs, controls: model.controls })
  const run = async <T>(fn: (context: unknown) => Promise<T>): Promise<T> => {
    runs += 1
    const ctx = new Ctx()
    const container: Container = { model, list: model.paragraphs, controls: model.controls }
    const selection = Object.assign(new RangeObj(ctx, container, null, 0, 0), {
      get paragraphs() { return new ParagraphCollection(ctx, container, () => model.selection) },
      get parentTableOrNullObject() {
        return model.selectionTable ? new TableObj(ctx, container, model.selectionTable) : new NullTableObj(ctx)
      },
      insertFootnote: (text: string) => ctx.queue(() => model.footnotes.push(text)),
    })
    Object.defineProperty(selection, 'text', { get: () => model.selectedText })
    const context = {
      sync: () => ctx.sync(),
      document: {
        body: body(ctx),
        sections: new SectionCollection(ctx, model),
        properties: new PropertiesObj(ctx, model),
        getSelection: () => selection,
        getStyles: () => ({ getByNameOrNullObject: (name: string) => new StyleObj(ctx, model, name) }),
      },
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
    context: {
      requirements: { isSetSupported: (_set: string, version: string) => Number(version) <= model.apiVersion },
      document: {
        get url() { return model.url },
        settings: {
          get: (k: string) => model.settings.get(k),
          set: (k: string, v: unknown) => model.settings.set(k, v),
          saveAsync: (cb: (r: { status: string }) => void) => cb({ status: 'succeeded' }),
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
