// Queries on the fake presentation, and helpers that change the selection the way clicks do (without firing
// DocumentSelectionChanged: use clickShapes in install.ts for that).

import type { FakePresentation, FakeShape, FakeSlide } from './model.ts'

/** Every shape on a slide, groups before their members, back to front. */
export function allShapes(s: FakeSlide | FakeShape[]): FakeShape[] {
  const list = Array.isArray(s) ? s : s.shapes
  return list.flatMap((x) => [x, ...(x.members ? allShapes(x.members) : [])])
}

export function findSlide(p: FakePresentation, id: string): FakeSlide | undefined {
  return p.slides.find((s) => s.id === id)
}

export function activeSlide(p: FakePresentation): FakeSlide | undefined {
  const id = p.selection.slideIds[0]
  return id === undefined ? undefined : findSlide(p, id)
}

/** A shape by id, on one slide or anywhere in the deck. */
export function findShape(p: FakePresentation, id: string, slideId?: string): FakeShape | undefined {
  const slides = slideId === undefined ? p.slides : p.slides.filter((s) => s.id === slideId)
  for (const s of slides) {
    const found = allShapes(s).find((x) => x.id === id)
    if (found) return found
  }
  return undefined
}

export interface Location {
  slide: FakeSlide
  /** The list that holds the shape: the slide's shapes or a group's members. */
  list: FakeShape[]
  parent: FakeShape | null
}

export function locate(p: FakePresentation, target: FakeShape): Location | null {
  const search = (slide: FakeSlide, list: FakeShape[], parent: FakeShape | null): Location | null => {
    if (list.includes(target)) return { slide, list, parent }
    for (const x of list) {
      const found = x.members ? search(slide, x.members, x) : null
      if (found) return found
    }
    return null
  }
  for (const s of p.slides) {
    const found = search(s, s.shapes, null)
    if (found) return found
  }
  return null
}

export function textOf(s: FakeShape): string {
  return s.textFrame ? s.textFrame.runs.map((r) => r.text).join('') : ''
}

/** Shapes carrying a tag (any value, or the given value), across the deck or on one slide. */
export function shapesTagged(p: FakePresentation | FakeSlide, key: string, value?: string): FakeShape[] {
  const slides = 'slides' in p ? p.slides : [p]
  return slides.flatMap((s) => allShapes(s)).filter((x) => x.tags.has(key.toUpperCase()) && (value === undefined || x.tags.get(key.toUpperCase()) === value))
}

export interface Bounds {
  left: number
  top: number
  width: number
  height: number
}

/** A shape's frame; for a group, the union of its members' frames, as PowerPoint reports it. */
export function boundsOf(s: FakeShape): Bounds {
  if (!s.members?.length) return { left: s.left, top: s.top, width: s.width, height: s.height }
  const bs = s.members.map(boundsOf)
  const left = Math.min(...bs.map((b) => b.left))
  const top = Math.min(...bs.map((b) => b.top))
  const right = Math.max(...bs.map((b) => b.left + b.width))
  const bottom = Math.max(...bs.map((b) => b.top + b.height))
  return { left, top, width: right - left, height: bottom - top }
}

/** Sets part of a shape's frame. A group moves or scales its members, as PowerPoint does. */
export function setBounds(s: FakeShape, next: Partial<Bounds>): void {
  if (!s.members?.length) {
    Object.assign(s, next)
    return
  }
  const b = boundsOf(s)
  const fx = next.width !== undefined && b.width ? next.width / b.width : 1
  const fy = next.height !== undefined && b.height ? next.height / b.height : 1
  const dx = (next.left ?? b.left) - b.left
  const dy = (next.top ?? b.top) - b.top
  const transform = (m: FakeShape): void => {
    if (m.members?.length) {
      m.members.forEach(transform)
      return
    }
    m.left = b.left + dx + (m.left - b.left) * fx
    m.top = b.top + dy + (m.top - b.top) * fy
    m.width *= fx
    m.height *= fy
  }
  s.members.forEach(transform)
}

// ---------------------------------------------------------------------------------------------- selection

/** Selects shapes on a slide (the active slide by default) without firing DocumentSelectionChanged. */
export function select(p: FakePresentation, shapeIds: string[], slideId?: string): void {
  const id = slideId ?? p.selection.slideIds[0] ?? p.slides[0]?.id
  p.selection = { slideIds: id === undefined ? [] : [id], shapeIds: [...shapeIds], textRange: null }
}

/** Selects slides; the first becomes the active slide. */
export function selectSlides(p: FakePresentation, slideIds: string[]): void {
  p.selection = { slideIds: [...slideIds], shapeIds: [], textRange: null }
}

/** Puts the cursor, or a text selection, in a shape on the active slide. */
export function selectText(p: FakePresentation, shapeId: string, start = 0, length = 0): void {
  const s = p.slides.find((x) => allShapes(x).some((y) => y.id === shapeId))
  p.selection = { slideIds: s ? [s.id] : p.selection.slideIds, shapeIds: [shapeId], textRange: { shapeId, start, length } }
}

/** The selected shapes on the active slide, in z-order (PowerPoint does not report the order of selection). */
export function selectedShapes(p: FakePresentation): FakeShape[] {
  const s = activeSlide(p)
  if (!s) return []
  const ids = new Set(p.selection.shapeIds)
  return allShapes(s).filter((x) => ids.has(x.id))
}
