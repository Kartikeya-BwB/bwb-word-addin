// Fake fields, comments, tracked changes and bookmarks.

import { Ctx, FakeCollection, FakeObject } from './base.ts'
import type { FakeBookmark, FakeModel } from './model.ts'

export class FieldObj extends FakeObject {
  constructor(context: Ctx, private readonly data: { code: string; updates: number }) {
    super(context)
  }
  get code(): string { this.need('code'); return this.data.code }
  updateResult(): void { this.context.queue(() => (this.data.updates += 1)) }
}

export class FieldCollection extends FakeCollection<FieldObj> {
  constructor(context: Ctx, private readonly model: FakeModel) {
    super(context)
  }
  protected make(): FieldObj[] {
    return this.model.fields.map((f) => new FieldObj(this.context, f))
  }
}

/** Inserts a field after a range, as Range.insertField does. The model records its code. */
export function insertField(context: Ctx, model: FakeModel, type: string, text: string): FieldObj {
  const data = { code: ` ${type.toUpperCase()} ${text} `, updates: 0 }
  context.queue(() => model.fields.push(data))
  return new FieldObj(context, data)
}

export class CommentObj extends FakeObject {
  constructor(context: Ctx, private readonly model: FakeModel) {
    super(context)
  }
  delete(): void { this.context.queue(() => (this.model.comments = Math.max(0, this.model.comments - 1))) }
}

export class CommentCollection extends FakeCollection<CommentObj> {
  constructor(context: Ctx, private readonly model: FakeModel) {
    super(context)
  }
  protected make(): CommentObj[] {
    return Array.from({ length: this.model.comments }, () => new CommentObj(this.context, this.model))
  }
}

/** Range.insertComment: counts the comment and keeps its text. */
export function insertComment(context: Ctx, model: FakeModel, text: string): CommentObj {
  context.queue(() => {
    model.comments += 1
    model.notes = [...(model.notes ?? []), text]
  })
  return new CommentObj(context, model)
}

class TrackedChangeObj extends FakeObject {}

export class TrackedChangeCollection extends FakeCollection<TrackedChangeObj> {
  constructor(context: Ctx, private readonly model: FakeModel) {
    super(context)
  }
  protected make(): TrackedChangeObj[] {
    return Array.from({ length: this.model.trackedChanges }, () => new TrackedChangeObj(this.context))
  }
  acceptAll(): void { this.settle('acceptAll') }
  rejectAll(): void { this.settle('rejectAll') }
  private settle(action: string): void {
    this.context.queue(() => {
      this.model.trackedChanges = 0
      this.model.reviewActions = [...(this.model.reviewActions ?? []), action]
    })
  }
}

/** The desktop-only document calls that act on every revision or comment at once. */
export function documentReview(context: Ctx, model: FakeModel): Record<'acceptAllRevisions' | 'rejectAllRevisions' | 'deleteAllComments', () => void> {
  const log = (action: string, change: () => void): (() => void) => () =>
    context.queue(() => {
      change()
      model.reviewActions = [...(model.reviewActions ?? []), action]
    })
  return {
    acceptAllRevisions: log('acceptAllRevisions', () => (model.trackedChanges = 0)),
    rejectAllRevisions: log('rejectAllRevisions', () => (model.trackedChanges = 0)),
    deleteAllComments: log('deleteAllComments', () => (model.comments = 0)),
  }
}

export function bookmarksOf(model: FakeModel): Map<string, FakeBookmark> {
  model.bookmarks ??= new Map()
  return model.bookmarks
}
