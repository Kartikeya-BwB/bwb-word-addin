// Brand: fonts other than Candara and Aptos, text, fill and outline colours outside the palette, lime on light
// backgrounds and text below the slide minimums (10 pt, or 8 pt for sources and footnotes).

import { SLIDE_TYPE, TAGS } from '../../brand/slideGrid.ts'
import type { ColourName } from '../../brand/tokens.ts'
import { ALLOWED_FONTS, COLOURS, FONTS } from '../../brand/tokens.ts'
import { SHAPE_COLOURS, TEXT_COLOURS, colourLabel, isLightBackground, isPaletteColour, nearestColour } from './colour.ts'
import { hasTag, hasText, isBwbFurniture, isFooterShape, isPicture, isSourceShape, isTitleShape } from './shapes.ts'
import { mergeSpans, points } from './text.ts'
import type { DeckFinding, DeckSnapshot, ShapeInfo, SlideInfo, Span } from './types.ts'
import { batchOf, foldFindings, onShape, plural, runsOf, shapeExcerpt, shapeKey, slideCount } from './util.ts'

/** More findings than this of one kind (one font, one colour) fold into one finding for the deck. */
const FOLD = 15

/** Theme font references such as +mj-lt, which PowerPoint resolves itself. */
const THEME_FONT = /^\+m[jn]-/i

/** Colours for fills and outlines: lime is for accents on navy only, so it is never offered as a replacement. */
const FILL_COLOURS: ColourName[] = SHAPE_COLOURS.filter((c) => c !== 'lime')

/** The stretches of a shape's text with something to see: not blank. */
function visibleRuns(shape: ShapeInfo): ReturnType<typeof runsOf> {
  const text = shape.text ?? ''
  return runsOf(shape).filter((r) => text.slice(r.start, r.start + r.length).trim() !== '')
}

function base(slide: SlideInfo, shape: ShapeInfo, span?: Span): Pick<DeckFinding, 'slide' | 'location' | 'excerpt'> {
  return { slide: slide.index + 1, location: onShape(slide, shape), excerpt: shapeExcerpt(shape, span) }
}

function fontFindings(snapshot: DeckSnapshot): DeckFinding[] {
  const out: DeckFinding[] = []
  for (const slide of snapshot.slides) {
    for (const shape of slide.shapes) {
      if (!hasText(shape) || isPicture(shape)) continue
      const byFont = new Map<string, Span[]>()
      for (const run of visibleRuns(shape)) {
        const name = run.font.name?.trim()
        if (!name || THEME_FONT.test(name) || ALLOWED_FONTS.has(name.toLowerCase())) continue
        byFont.set(name, [...(byFont.get(name) ?? []), { start: run.start, length: run.length }])
      }
      for (const [font, spans] of byFont) {
        const target = isTitleShape(shape) ? FONTS.heading : FONTS.body
        out.push({
          id: `brand:font:${shapeKey(slide, shape)}:${font}`, rule: 'brand/font', group: 'brand', severity: 'minor',
          title: `Text in ${font}`, detail: 'BwB slides use Candara for titles and Aptos for text.', ...base(slide, shape, spans[0]),
          fix: { kind: 'font', slideId: slide.id, path: shape.path, expectedText: shape.text ?? '', spans: mergeSpans(spans, shape.text ?? ''), font: target },
          fixLabel: `Change to ${target}`,
        })
      }
    }
  }
  return foldFindings(out, (f) => f.title, FOLD, (alike, title) => {
    const font = title.replace(/^Text in /, '')
    const fix = batchOf(alike)
    const first = alike[0]
    return {
      id: `brand:font-deck:${font}`, rule: 'brand/deck-font', group: 'brand', severity: 'minor',
      title: `${font} in ${plural(alike.length, 'shape')} on ${plural(slideCount(alike), 'slide')}`,
      detail: 'The deck’s theme fonts are probably not BwB’s. Start from the BwB template, or change every shape here: titles to Candara, the rest to Aptos.',
      ...(first?.slide !== undefined ? { slide: first.slide } : {}), ...(first?.location ? { location: first.location } : {}),
      ...(fix ? { fix, fixLabel: 'Change them all to the BwB fonts' } : {}),
    }
  })
}

type ColourKind = 'text' | 'fill' | 'line'

const KIND_WORD: Record<ColourKind, string> = { text: 'Text colour', fill: 'Fill', line: 'Outline' }

function colourFinding(slide: SlideInfo, shape: ShapeInfo, kind: ColourKind, hex: string, spans: Span[]): DeckFinding {
  const replacement = nearestColour(hex, kind === 'text' ? TEXT_COLOURS : FILL_COLOURS)
  const name = colourLabel(replacement.hex)
  const ref = { slideId: slide.id, path: shape.path }
  return {
    id: `brand:colour:${kind}:${shapeKey(slide, shape)}:${hex}`, rule: 'brand/off-palette', group: 'brand', severity: 'minor',
    title: `${KIND_WORD[kind]} #${hex} is not in the BwB palette`,
    detail: kind === 'text'
      ? 'Use navy, charcoal or black for text; white on navy; green only at 18 pt and above.'
      : 'Use the BwB palette: navy, green, charcoal, the greys and off-white; lime only on navy.',
    ...base(slide, shape, spans[0]),
    fix: kind === 'text'
      ? { ...ref, kind: 'text-colour', expectedText: shape.text ?? '', spans: mergeSpans(spans, shape.text ?? ''), colour: replacement.hex }
      : { ...ref, kind: kind === 'fill' ? 'fill-colour' : 'line-colour', colour: replacement.hex },
    fixLabel: `Change to ${name}`,
  }
}

function colourFindings(snapshot: DeckSnapshot): DeckFinding[] {
  const out: DeckFinding[] = []
  for (const slide of snapshot.slides) {
    for (const shape of slide.shapes) {
      if (isPicture(shape)) continue
      const fill = shape.fill?.type === 'Solid' ? shape.fill.colour : null
      const byColour = new Map<string, Span[]>()
      const lime: Span[] = []
      for (const run of visibleRuns(shape)) {
        const c = run.font.colour
        if (!c) continue
        const span = { start: run.start, length: run.length }
        if (c === COLOURS.lime && fill !== null && isLightBackground(fill)) lime.push(span)
        else if (!isPaletteColour(c)) byColour.set(c, [...(byColour.get(c) ?? []), span])
      }
      for (const [hex, spans] of byColour) out.push(colourFinding(slide, shape, 'text', hex, spans))
      if (lime.length) {
        out.push({
          id: `brand:lime:${shapeKey(slide, shape)}`, rule: 'brand/lime-on-light', group: 'brand', severity: 'minor',
          title: 'Lime text on a light background', detail: 'Lime is for navy backgrounds only; it fails contrast on white and the greys.',
          ...base(slide, shape, lime[0]),
          fix: { kind: 'text-colour', slideId: slide.id, path: shape.path, expectedText: shape.text ?? '', spans: mergeSpans(lime, shape.text ?? ''), colour: COLOURS.navy },
          fixLabel: 'Change to Navy',
        })
      }
      // Red, amber and green status dots are outside the palette on purpose.
      if (hasTag(shape, TAGS.rag)) continue
      if (fill !== null && !isPaletteColour(fill)) out.push(colourFinding(slide, shape, 'fill', fill, []))
      const line = shape.line?.visible ? shape.line.colour : null
      if (line !== null && !isPaletteColour(line)) out.push(colourFinding(slide, shape, 'line', line, []))
    }
  }
  return foldFindings(out, (f) => (f.rule === 'brand/off-palette' ? f.title : null), FOLD, (alike, title) => {
    const fix = batchOf(alike)
    const first = alike[0]
    return {
      id: `brand:colour-deck:${title}`, rule: 'brand/deck-colour', group: 'brand', severity: 'minor',
      title: `${title.replace(/ is not in the BwB palette$/, '')} in ${plural(alike.length, 'shape')} on ${plural(slideCount(alike), 'slide')}`,
      detail: 'The colour is not in the BwB palette. The deck’s theme colours are probably not BwB’s: start from the BwB template, or change every shape here.',
      ...(first?.slide !== undefined ? { slide: first.slide } : {}), ...(first?.location ? { location: first.location } : {}),
      ...(fix ? { fix, fixLabel: first?.fixLabel ? first.fixLabel.replace(/^Change to/, 'Change them all to') : 'Change them all' } : {}),
    }
  })
}

function smallTextFindings(snapshot: DeckSnapshot): DeckFinding[] {
  const out: DeckFinding[] = []
  for (const slide of snapshot.slides) {
    for (const shape of slide.shapes) {
      if (!hasText(shape) || isPicture(shape)) continue
      const relaxed = isSourceShape(shape) || isFooterShape(shape) || isBwbFurniture(shape)
      const min = relaxed ? SLIDE_TYPE.minAny : SLIDE_TYPE.minBody
      const small = visibleRuns(shape).filter((r) => r.font.size !== null && r.font.size < min)
      if (!small.length) continue
      const smallest = Math.min(...small.map((r) => r.font.size ?? min))
      const spans = mergeSpans(small.map((r) => ({ start: r.start, length: r.length })), shape.text ?? '')
      out.push({
        id: `brand:small:${shapeKey(slide, shape)}`, rule: 'brand/small-text', group: 'brand', severity: 'minor',
        title: `Text at ${points(smallest)}`,
        detail: relaxed
          ? `Sources and footnotes go down to ${SLIDE_TYPE.minAny} pt, no further.`
          : `Slide text stays at ${SLIDE_TYPE.minBody} pt or more. Only sources and footnotes go down to ${SLIDE_TYPE.minAny} pt.`,
        ...base(slide, shape, spans[0]),
        fix: { kind: 'font-size', slideId: slide.id, path: shape.path, expectedText: shape.text ?? '', spans, size: min },
        fixLabel: `Set to ${min} pt`,
      })
    }
  }
  return out
}

export function brandFindings(snapshot: DeckSnapshot): DeckFinding[] {
  if (!snapshot.can.formatting) return []
  return [...fontFindings(snapshot), ...colourFindings(snapshot), ...smallTextFindings(snapshot)]
}
