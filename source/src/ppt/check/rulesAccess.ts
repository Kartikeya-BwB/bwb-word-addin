// Readability and what travels with the file: text colour against what is behind it (WCAG AA), shapes hidden on
// slides, and links to places readers outside BwB cannot open.

import { TAGS } from '../../brand/slideGrid.ts'
import { COLOURS } from '../../brand/tokens.ts'
import { TEXT_COLOURS, colourLabel, contrastNeeded, contrastRatio, isLightBackground } from './colour.ts'
import { INTERNAL_LINK_DETAIL, addressesIn, internalLink } from './links.ts'
import { hasTag, hasText, isPicture } from './shapes.ts'
import { mergeSpans } from './text.ts'
import type { DeckFinding, DeckSnapshot, ShapeInfo, SlideInfo, Span } from './types.ts'
import { excerpt, onShape, onSlide, runsOf, shapeExcerpt, shapeKey } from './util.ts'
import { boxOf } from './rulesLayout.ts'

/** Shapes whose colours the check cannot see: pictures, charts, diagrams and the like. */
const OPAQUE_UNKNOWN = new Set(['Image', 'Graphic', 'Chart', 'SmartArt', 'Media', 'Table', 'Ole', 'ContentApp', 'Diagram', 'Model3D', 'Unsupported'])

/** 'navy' for a palette colour in a sentence; '#123456' otherwise. */
const colourWord = (hex: string): string => {
  const label = colourLabel(hex)
  return label.startsWith('#') ? label : label.toLowerCase()
}

const inside = (s: ShapeInfo, x: number, y: number): boolean => x >= s.left && x <= s.left + s.width && y >= s.top && y <= s.top + s.height

/**
 * The colour behind a shape's text: its own solid fill, else the nearest filled shape underneath it, else the slide
 * background. Null when that cannot be known (a picture underneath, a see-through fill, an unreadable background).
 */
export function colourBehind(slide: SlideInfo, index: number): string | null {
  const shape = slide.shapes[index]
  if (!shape) return null
  const solid = (s: ShapeInfo): string | null | undefined => {
    if (s.fill === null) return undefined
    if (s.fill.type !== 'Solid') return s.fill.type === 'NoFill' ? undefined : null
    return (s.fill.transparency ?? 0) > 0.1 ? null : s.fill.colour
  }
  const own = solid(shape)
  if (own !== undefined) return own
  const x = shape.left + shape.width / 2
  const y = shape.top + shape.height / 2
  for (let i = index - 1; i >= 0; i -= 1) {
    const below = slide.shapes[i]
    if (!below || below.visible === false || below.type === 'Group' || below.type === 'Line' || !inside(below, x, y)) continue
    if (OPAQUE_UNKNOWN.has(below.type) || isPicture(below)) return null
    const fill = solid(below)
    if (fill !== undefined) return fill
  }
  return slide.background
}

function contrastFindings(snapshot: DeckSnapshot): DeckFinding[] {
  if (!snapshot.can.formatting) return []
  const out: DeckFinding[] = []
  for (const slide of snapshot.slides) {
    slide.shapes.forEach((shape, index) => {
      if (!hasText(shape) || isPicture(shape) || hasTag(shape, TAGS.rag)) return
      const back = colourBehind(slide, index)
      if (!back) return
      const text = shape.text ?? ''
      const byColour = new Map<string, { spans: Span[]; ratio: number; need: number }>()
      for (const run of runsOf(shape)) {
        const fore = run.font.colour
        if (!fore || !text.slice(run.start, run.start + run.length).trim()) continue
        // Lime on a light colour has its own finding.
        if (fore === COLOURS.lime && isLightBackground(back)) continue
        const ratio = contrastRatio(fore, back)
        const need = contrastNeeded(run.font.size ?? shape.font?.size ?? null, run.font.bold ?? shape.font?.bold ?? null)
        if (ratio >= need) continue
        const entry = byColour.get(fore) ?? { spans: [], ratio, need }
        entry.spans.push({ start: run.start, length: run.length })
        entry.ratio = Math.min(entry.ratio, ratio)
        entry.need = Math.max(entry.need, need)
        byColour.set(fore, entry)
      }
      for (const [fore, { spans, ratio, need }] of byColour) {
        const best = TEXT_COLOURS.map((key) => ({ key, hex: COLOURS[key], ratio: contrastRatio(COLOURS[key], back) })).sort((a, b) => b.ratio - a.ratio)[0]
        const fixable = best !== undefined && best.ratio >= need
        out.push({
          id: `brand:contrast:${shapeKey(slide, shape)}:${fore}`, rule: 'brand/contrast', group: 'brand', severity: 'minor',
          title: `Low contrast (${Math.floor(ratio * 10) / 10}:1)`,
          detail: `${colourWord(fore).replace(/^./, (c) => c.toUpperCase())} text on ${colourWord(back)} needs a contrast of at least ${need}:1 to read on screens, projectors and paper (WCAG AA). Use navy or black on light colours and white on navy.`,
          excerpt: shapeExcerpt(shape, spans[0]), slide: slide.index + 1, location: onShape(slide, shape),
          ...(fixable && best
            ? {
                fix: { kind: 'text-colour' as const, slideId: slide.id, path: shape.path, expectedText: text, spans: mergeSpans(spans, text), colour: best.hex },
                fixLabel: `Change to ${colourLabel(best.hex)}`,
              }
            : {}),
        })
      }
    })
  }
  return out
}

/** Shapes hidden on their slides. Run on the whole snapshot: the other rules see only the shapes that show. */
export function hiddenFindings(snapshot: DeckSnapshot): DeckFinding[] {
  const out: DeckFinding[] = []
  for (const slide of snapshot.slides) {
    const hidden = new Set(slide.shapes.filter((s) => s.visible === false).map((s) => s.id))
    for (const shape of slide.shapes) {
      // Report a hidden group once, not each shape in it.
      if (!hidden.has(shape.id) || shape.path.slice(0, -1).some((id) => hidden.has(id))) continue
      out.push({
        id: `send:hidden:${shapeKey(slide, shape)}`, rule: 'send/hidden-shape', group: 'send', severity: 'minor',
        title: 'Hidden shape',
        detail: 'It does not show on the slide, but it stays in the file, where anyone can unhide it. Delete it if it is not needed.',
        excerpt: shapeExcerpt(shape), slide: slide.index + 1, location: onShape(slide, shape),
        fix: { kind: 'delete-shape', slideId: slide.id, path: shape.path, from: boxOf(shape) }, fixLabel: 'Delete it',
      })
    }
  }
  return out
}

function linkFindings(snapshot: DeckSnapshot): DeckFinding[] {
  const out: DeckFinding[] = []
  for (const slide of snapshot.slides) {
    const seen = new Set<string>()
    const add = (address: string, shape: ShapeInfo | undefined): void => {
      const kind = internalLink(address)
      if (!kind || seen.has(address)) return
      seen.add(address)
      out.push({
        id: `send:link:${slide.id}:${address}`, rule: 'send/internal-link', group: 'send', severity: 'minor',
        title: 'Link to an internal location', detail: INTERNAL_LINK_DETAIL[kind], excerpt: excerpt(address, 0, Math.min(address.length, 80), 0),
        slide: slide.index + 1, location: shape ? onShape(slide, shape) : onSlide(slide),
      })
    }
    for (const link of slide.links) add(link.address, link.shapeId ? slide.shapes.find((s) => s.id === link.shapeId) : undefined)
    for (const shape of slide.shapes) if (hasText(shape)) for (const address of addressesIn(shape.text ?? '')) add(address, shape)
  }
  return out
}

export function accessFindings(snapshot: DeckSnapshot): DeckFinding[] {
  return [...contrastFindings(snapshot), ...linkFindings(snapshot)]
}

