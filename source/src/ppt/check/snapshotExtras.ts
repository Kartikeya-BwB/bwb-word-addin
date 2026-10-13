// What the Pre-QA check reads beyond the shapes, a sync or two each for the whole deck: each slide's background colour
// (API 1.10), its hyperlinks (1.6) and the superscript footnote markers in its text (1.8).

import { FOOTNOTES_TAG, markerCandidates } from '../slides/footnotes.ts'
import { syncEach } from '../slides/office.ts'
import { hasTag, isBwbFurniture } from './shapes.ts'
import { hexColour } from './text.ts'
import type { ShapeInfo, SlideInfo, SlideLink, Span } from './types.ts'

export interface SlideRef {
  slide: PowerPoint.Slide
  info: SlideInfo
}

export interface TextRef {
  range: PowerPoint.TextRange | null
  info: ShapeInfo
}

/** Number candidates read for superscript in one deck, at most. */
const MAX_CANDIDATES = 3000

interface BackgroundProbe {
  ref: SlideRef
  follows: [PowerPoint.SlideBackground, PowerPoint.SlideLayoutBackground] | null
  solids: PowerPoint.SlideBackgroundSolidFill[]
}

/** The colour each slide shows behind its shapes: its own background, its layout's, or its master's, when solid. */
export async function readBackgrounds(context: PowerPoint.RequestContext, slides: SlideRef[]): Promise<void> {
  const probes: BackgroundProbe[] = slides.map((ref) => ({ ref, follows: null, solids: [] }))
  const failed = await syncEach(context, probes, (p) => {
    const own = p.ref.slide.background
    const layout = p.ref.slide.layout.background
    own.load('isMasterBackgroundFollowed')
    layout.load('isMasterBackgroundFollowed')
    p.follows = [own, layout]
    p.solids = [own.fill, layout.fill, p.ref.slide.slideMaster.background.fill].map((fill) => {
      const solid = fill.getSolidFillOrNullObject()
      solid.load('color')
      return solid
    })
  })
  for (const p of probes) {
    if (failed.has(p) || !p.follows) continue
    const [own, layout] = p.follows
    const solid = p.solids[!own.isMasterBackgroundFollowed ? 0 : !layout.isMasterBackgroundFollowed ? 1 : 2]
    p.ref.info.background = solid && !solid.isNullObject ? hexColour(solid.color) : null
  }
}

interface LinkProbe {
  link: PowerPoint.Hyperlink
  info: SlideLink
  shape: PowerPoint.Shape | null
}

/** Each slide's hyperlink addresses, and with API 1.10 the shape each link is on. */
export async function readLinks(context: PowerPoint.RequestContext, slides: SlideRef[], linkedShapes: boolean): Promise<void> {
  const probes = slides.map((ref) => ({ ref, links: null as PowerPoint.HyperlinkCollection | null }))
  const failed = await syncEach(context, probes, (p) => {
    p.links = p.ref.slide.hyperlinks
    p.links.load('items/address')
  })
  const found: LinkProbe[] = []
  for (const p of probes) {
    if (failed.has(p) || !p.links) continue
    for (const link of p.links.items) {
      const info: SlideLink = { address: link.address ?? '', shapeId: null }
      p.ref.info.links.push(info)
      found.push({ link, info, shape: null })
    }
  }
  if (!linkedShapes || !found.length) return
  const unread = await syncEach(context, found, (f) => {
    f.shape = f.link.getLinkedShapeOrNullObject()
    f.shape.load('id')
  })
  for (const f of found) if (!unread.has(f) && f.shape && !f.shape.isNullObject) f.info.shapeId = f.shape.id
}

interface MarkerProbe {
  ref: TextRef
  span: Span
  font: PowerPoint.ShapeFont | null
}

/** Superscript numbers in the text, read one candidate at a time. Returns false when there were too many to read. */
export async function readMarkers(context: PowerPoint.RequestContext, shapes: TextRef[]): Promise<boolean> {
  const probes: MarkerProbe[] = []
  for (const ref of shapes) {
    const text = ref.info.text
    if (!ref.range || !text || hasTag(ref.info, FOOTNOTES_TAG) || isBwbFurniture(ref.info)) continue
    for (const span of markerCandidates(text)) probes.push({ ref, span, font: null })
  }
  const batch = probes.slice(0, MAX_CANDIDATES)
  const failed = await syncEach(context, batch, (p) => {
    p.font = p.ref.range ? p.ref.range.getSubstring(p.span.start, p.span.length).font : null
    p.font?.load('superscript')
  })
  for (const p of batch) {
    if (failed.has(p) || !p.font || p.font.superscript !== true) continue
    const { start, length } = p.span
    p.ref.info.markers.push({ start, length, number: Number((p.ref.info.text ?? '').slice(start, start + length)) })
  }
  return probes.length <= MAX_CANDIDATES
}
