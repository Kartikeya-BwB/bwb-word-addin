// The object getSelection() returns. Tests describe the selection with model.selection (the paragraphs it touches),
// and optionally model.selectionRange (a span inside one paragraph) and model.selectedPictures.

import type { Ctx } from './base.ts'
import type { FakeModel, FakeParagraph } from './model.ts'
import type { Container } from './objects.ts'
import { FontObj, ParagraphCollection, RangeObj } from './objects.ts'
import { PictureCollection } from './pictures.ts'
import { NullTableObj, TableObj } from './tables.ts'

export class SelectionObj extends RangeObj {
  constructor(context: Ctx, container: Container, private readonly model: FakeModel) {
    const span = model.selectionRange
    const first: FakeParagraph | null = span?.paragraph ?? model.selection[0] ?? null
    // Without a span, the cursor sits at the end of the first selected paragraph.
    const at = first?.text.length ?? 0
    super(context, container, first, span?.start ?? at, span?.end ?? at)
    this.selectionOf = model
  }
  private covered(): FakeParagraph[] {
    return this.model.selectionRange ? [this.model.selectionRange.paragraph] : this.model.selection
  }
  override get text(): string {
    const span = this.model.selectionRange
    return span ? span.paragraph.text.slice(span.start, span.end) : this.model.selectedText
  }
  override get font(): FontObj {
    return new FontObj(this.context, () => this.covered().map((p) => p.font))
  }
  override set case(v: string) {
    this.context.queue(() => {
      const span = this.model.selectionRange
      const change = (s: string): string => (v === 'Upper' ? s.toUpperCase() : v === 'Lower' ? s.toLowerCase() : s)
      if (span) {
        const p = span.paragraph
        p.text = p.text.slice(0, span.start) + change(p.text.slice(span.start, span.end)) + p.text.slice(span.end)
      } else this.model.selection.forEach((p) => (p.text = change(p.text)))
    })
  }
  get paragraphs(): ParagraphCollection {
    return new ParagraphCollection(this.context, this.container, () => this.covered())
  }
  get parentTableOrNullObject(): TableObj | NullTableObj {
    return this.model.selectionTable ? new TableObj(this.context, this.container, this.model.selectionTable) : new NullTableObj(this.context)
  }
  get inlinePictures(): PictureCollection {
    return new PictureCollection(this.context, this.container, () => this.model.selectedPictures ?? [])
  }
  insertFootnote(text: string): void {
    this.context.queue(() => this.model.footnotes.push(text))
  }
}
