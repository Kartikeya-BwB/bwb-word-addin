// Fake content controls (moved out of objects.ts so that file stays small).

import { Ctx, FakeCollection, FakeObject } from './base.ts'
import type { FakeControl } from './model.ts'
import type { Container } from './objects.ts'

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
