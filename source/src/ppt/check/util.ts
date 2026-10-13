// Small helpers the deck rules share: locations, formatting runs and folding many alike findings into one.

import { excerpt } from '../../check/util.ts'
import type { DeckFinding, DeckFix, DeckLocation, FontInfo, ShapeInfo, SlideInfo, Span } from './types.ts'

export { excerpt }

export function onSlide(slide: SlideInfo): DeckLocation {
  return { kind: 'slide', slideId: slide.id, slide: slide.index + 1 }
}

export function onShape(slide: SlideInfo, shape: ShapeInfo, start?: number, end?: number): DeckLocation {
  return {
    kind: 'shape', slideId: slide.id, slide: slide.index + 1, path: shape.path,
    ...(start !== undefined ? { start } : {}), ...(end !== undefined ? { end } : {}),
  }
}

/** A stable id for a finding on a shape. */
export function shapeKey(slide: SlideInfo, shape: ShapeInfo): string {
  return `${slide.id}:${shape.path.join('/')}`
}

/** The shape's text in uniform stretches: the pieces read for mixed text, else the whole text. */
export function runsOf(shape: ShapeInfo): Array<Span & { font: FontInfo }> {
  if (shape.pieces.length) return shape.pieces
  if (!shape.font || !shape.text) return []
  return [{ start: 0, length: shape.text.length, font: shape.font }]
}

/** A short excerpt of a shape's text around a span, on one line; the shape's name when it has no text. */
export function shapeExcerpt(shape: ShapeInfo, span?: Span): string {
  const text = shape.text ?? ''
  if (!text.trim()) return shape.name
  return span ? excerpt(text, span.start, span.start + span.length) : excerpt(text, 0, Math.min(text.length, 60), 0)
}

export function plural(n: number, one: string, many = `${one}s`): string {
  return `${n} ${n === 1 ? one : many}`
}

/**
 * Folds findings that share a key into one when there are more than `limit` of them, so a deck set in Calibri gives
 * one finding with one fix rather than hundreds. Findings without a key, or in small groups, pass through.
 */
export function foldFindings(
  findings: DeckFinding[],
  key: (f: DeckFinding) => string | null,
  limit: number,
  fold: (alike: DeckFinding[], key: string) => DeckFinding,
): DeckFinding[] {
  const groups = new Map<string, DeckFinding[]>()
  for (const f of findings) {
    const k = key(f)
    if (k !== null) groups.set(k, [...(groups.get(k) ?? []), f])
  }
  const folded = new Map<string, DeckFinding>()
  for (const [k, alike] of groups) if (alike.length > limit) folded.set(k, fold(alike, k))
  const out: DeckFinding[] = []
  const emitted = new Set<string>()
  for (const f of findings) {
    const k = key(f)
    const replacement = k === null ? undefined : folded.get(k)
    if (!replacement || k === null) {
      out.push(f)
    } else if (!emitted.has(k)) {
      emitted.add(k)
      out.push(replacement)
    }
  }
  return out
}

/** One fix that applies every finding's fix. */
export function batchOf(findings: DeckFinding[]): DeckFix | undefined {
  const fixes = findings.flatMap((f) => (f.fix ? [f.fix] : []))
  return fixes.length ? { kind: 'batch', fixes } : undefined
}

/** How many different slides the findings are on. */
export function slideCount(findings: DeckFinding[]): number {
  return new Set(findings.map((f) => f.slide).filter((n) => n !== undefined)).size
}
