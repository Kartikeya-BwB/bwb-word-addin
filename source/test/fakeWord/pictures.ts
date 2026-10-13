// Fake inline pictures: size, proportions, alt text and the paragraph that holds each one.

import { Ctx, FakeCollection, FakeObject } from './base.ts'
import type { FakePicture } from './model.ts'
import { createModel, para } from './model.ts'
import type { Container } from './objects.ts'
import { ParagraphObj } from './objects.ts'

export class PictureObj extends FakeObject {
  constructor(context: Ctx, private readonly container: Container | null, readonly data: FakePicture) {
    super(context)
  }
  get altTextTitle(): string { this.need('altTextTitle'); return this.data.altTextTitle }
  set altTextTitle(v: string) { this.context.queue(() => (this.data.altTextTitle = v)) }
  get altTextDescription(): string { this.need('altTextDescription'); return this.data.altTextDescription }
  set altTextDescription(v: string) { this.context.queue(() => (this.data.altTextDescription = v)) }
  get width(): number { this.need('width'); return this.data.width ?? 300 }
  get height(): number { this.need('height'); return this.data.height ?? 200 }
  get lockAspectRatio(): boolean { this.need('lockAspectRatio'); return this.data.lockAspectRatio ?? true }
  set lockAspectRatio(v: boolean) { this.context.queue(() => (this.data.lockAspectRatio = v)) }
  // Word scales the other side when the proportions are locked, as it does for a picture dragged by a corner.
  set width(v: number) {
    this.context.queue(() => {
      const { width = 300, height = 200, lockAspectRatio = true } = this.data
      if (lockAspectRatio && width > 0) this.data.height = (height * v) / width
      this.data.width = v
    })
  }
  set height(v: number) {
    this.context.queue(() => {
      const { width = 300, height = 200, lockAspectRatio = true } = this.data
      if (lockAspectRatio && height > 0) this.data.width = (width * v) / height
      this.data.height = v
    })
  }
  get paragraph(): ParagraphObj {
    // Pictures in header parts and table cells are not tied to a body paragraph.
    const container: Container = this.container ?? { model: createModel([]), list: [], controls: [] }
    return new ParagraphObj(this.context, container, this.data.paragraph ?? para(''))
  }
  select(): void {
    this.context.queue(() => this.container?.model.selected.push(`picture:${this.data.altTextTitle}`))
  }
}

export class PictureCollection extends FakeCollection<PictureObj> {
  constructor(context: Ctx, private readonly container: Container | null, private readonly list: () => FakePicture[]) {
    super(context)
  }
  protected make(): PictureObj[] {
    return this.list().map((p) => new PictureObj(this.context, this.container, p))
  }
  getFirstOrNullObject(): PictureObj | { isNullObject: true } {
    const first = this.list()[0]
    return first ? new PictureObj(this.context, this.container, first) : { isNullObject: true }
  }
}
