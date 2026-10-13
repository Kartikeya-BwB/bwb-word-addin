// The BwB slide elements as plain layouts: shapes with positions, colours and text, built from the slide grid and
// the slide type sizes. render.ts turns a layout into PowerPoint shapes; the tests check layouts without PowerPoint.

import { ICON_SIZES_CM } from '../brand/icons.ts'
import { CONTENT_AREA, GUTTER, MARGINS, SLIDE, SLIDE_TYPE, SOURCE_AREA, columnEdges } from '../brand/slideGrid.ts'
import { COLOURS } from '../brand/tokens.ts'
import type { Rect } from '../geometry/index.ts'
import { cmToPt } from '../geometry/index.ts'

export type ElementKind =
  | 'callout' | 'key-figures' | 'quote' | 'process' | 'matrix' | 'timeline' | 'step-marker' | 'source' | 'table' | 'note' | 'sticker'
  | 'rag' | 'harvey' | 'status'

export interface FontSpec {
  name: string
  size: number
  /** Hex without '#'. */
  colour: string
  bold?: boolean
  italic?: boolean
}

export interface TextSpec {
  paragraphs: Array<{ text: string; font: FontSpec }>
  align: 'Left' | 'Center' | 'Right'
  vertical: 'Top' | 'Middle' | 'Bottom'
  margins: { left: number; right: number; top: number; bottom: number }
  wrap: boolean
}

export interface ShapeSpec {
  kind: 'geometric' | 'text' | 'line'
  geometry: 'Rectangle' | 'Ellipse' | 'Chevron' | 'HomePlate'
  name: string
  left: number
  top: number
  width: number
  height: number
  /** Degrees clockwise; needs PowerPointApi 1.10. */
  rotation?: number
  /** Hex without '#'; null for none. */
  fill: string | null
  line: { colour: string; weight: number } | null
  text?: TextSpec
}

export interface ElementSpec {
  kind: ElementKind
  name: string
  shapes: ShapeSpec[]
  /** Group the shapes when PowerPoint can (PowerPointApi 1.8). */
  group: boolean
}

/** Inner padding and small gaps: half the grid gutter. */
const PAD = GUTTER / 2
/** The coloured bar of a callout or quote, as wide as the 4 pt bar of the Word callout. */
const BAR = 4
/** PowerPoint's single line spacing is about 1.2 times the type size. */
const LINE = 1.2
/** A line of small text with room around it: axis labels. */
const LABEL = GUTTER * 1.5
/** The weight of the timeline's line, in points. */
const AXIS_WEIGHT = 2
/** Light grey rules, as on BwB tables. */
export const RULE = { colour: COLOURS.lightGrey, weight: 0.75 } as const

const APTOS = SLIDE_TYPE.body.font
const body = (over: Partial<FontSpec> = {}): FontSpec => ({ name: APTOS, size: SLIDE_TYPE.body.size, colour: COLOURS.navy, ...over })
const small = (over: Partial<FontSpec> = {}): FontSpec => ({ name: APTOS, size: SLIDE_TYPE.small.size, colour: COLOURS.navy, ...over })

const text = (paragraphs: TextSpec['paragraphs'], over: Partial<TextSpec> = {}): TextSpec => ({
  paragraphs, align: 'Left', vertical: 'Top', margins: { left: PAD, right: PAD, top: PAD / 2, bottom: PAD / 2 }, wrap: true, ...over,
})

const NO_MARGINS = { left: 0, right: 0, top: 0, bottom: 0 }

const rect = (name: string, r: Rect, fill: string | null, over: Partial<ShapeSpec> = {}): ShapeSpec => ({
  kind: 'geometric', geometry: 'Rectangle', name, ...r, fill, line: null, ...over,
})

/** Columns first to last of the BwB grid as one area, as tall as the content area. */
function columns(first: number, last: number): Rect {
  const cols = columnEdges()
  const left = cols[first]?.left ?? CONTENT_AREA.left
  const right = cols[last]?.right ?? CONTENT_AREA.left + CONTENT_AREA.width
  return { left, top: CONTENT_AREA.top, width: right - left, height: CONTENT_AREA.height }
}

const bottomOfContent = CONTENT_AREA.top + CONTENT_AREA.height

/** An off-white box with a green bar, a bold title and a message, along the bottom of the content area. */
export function calloutSpec(): ElementSpec {
  const height = 3 * SLIDE_TYPE.body.size * LINE + 2 * PAD
  const top = bottomOfContent - height
  return {
    kind: 'callout', name: 'BwB callout', group: true,
    shapes: [
      rect('BwB callout bar', { left: CONTENT_AREA.left, top, width: BAR, height }, COLOURS.green),
      rect('BwB callout text', { left: CONTENT_AREA.left + BAR, top, width: CONTENT_AREA.width - BAR, height }, COLOURS.offWhite, {
        text: text([{ text: 'Callout title', font: body({ bold: true }) }, { text: 'One or two sentences with the message.', font: body() }], {
          vertical: 'Middle', margins: { left: 2 * PAD, right: 2 * PAD, top: PAD / 2, bottom: PAD / 2 },
        }),
      }),
    ],
  }
}

/** Three navy panels across the content area, each four columns wide: a large figure (the first in lime) and a label. */
export function keyFiguresSpec(): ElementSpec {
  const height = SLIDE_TYPE.title.size * LINE + 2 * SLIDE_TYPE.small.size * LINE + 4 * PAD
  return {
    kind: 'key-figures', name: 'BwB key figures', group: true,
    shapes: [0, 1, 2].map((i) => rect(`BwB key figure ${i + 1}`, { ...columns(i * 4, i * 4 + 3), height }, COLOURS.navy, {
      text: text([
        { text: 'Figure', font: body({ size: SLIDE_TYPE.title.size, bold: true, colour: i === 0 ? COLOURS.lime : COLOURS.white }) },
        { text: 'What the figure measures', font: small({ colour: COLOURS.white }) },
      ], { vertical: 'Middle', margins: { left: 2 * PAD, right: 2 * PAD, top: PAD, bottom: PAD } }),
    })),
  }
}

/** A quotation in Candara italic navy beside a lime bar. The bar sits on a navy strip: lime appears only on navy. */
export function quoteSpec(): ElementSpec {
  const height = 2 * SLIDE_TYPE.title.size * LINE + SLIDE_TYPE.subtitle.size * LINE + 2 * PAD
  const { left, top, width } = CONTENT_AREA
  return {
    kind: 'quote', name: 'BwB quote', group: true,
    shapes: [
      rect('BwB quote strip', { left, top, width: PAD, height }, COLOURS.navy),
      rect('BwB quote bar', { left: left + (PAD - BAR) / 2, top, width: BAR, height }, COLOURS.lime),
      {
        kind: 'text', geometry: 'Rectangle', name: 'BwB quote text', left: left + PAD + GUTTER, top, width: width - PAD - GUTTER, height, fill: null, line: null,
        text: text([
          {
            text: '“The quotation, in the speaker’s own words.”',
            font: { name: SLIDE_TYPE.title.font, size: SLIDE_TYPE.title.size, colour: COLOURS.navy, italic: true },
          },
          { text: 'Name, role, organisation', font: { name: SLIDE_TYPE.subtitle.font, size: SLIDE_TYPE.subtitle.size, colour: SLIDE_TYPE.subtitle.colour } },
        ], { vertical: 'Middle', margins: NO_MARGINS }),
      },
    ],
  }
}

/** Three chevrons across the content area, nested half a gutter apart: two navy steps and a green outcome. */
export function processSpec(): ElementSpec {
  const height = 4 * GUTTER
  // PowerPoint's chevrons point by half their height; each nests in the notch of the one before.
  const overlap = height / 2 - PAD
  const width = (CONTENT_AREA.width + 2 * overlap) / 3
  const fills = [COLOURS.navy, COLOURS.navy, COLOURS.green]
  return {
    kind: 'process', name: 'BwB process', group: true,
    shapes: fills.map((fill, i) => ({
      kind: 'geometric', geometry: i === 0 ? 'HomePlate' : 'Chevron', name: `BwB process step ${i + 1}`,
      left: CONTENT_AREA.left + i * (width - overlap), top: CONTENT_AREA.top, width, height, fill, line: null,
      text: text([{ text: `Step ${i + 1}`, font: body({ colour: COLOURS.white, bold: true }) }], {
        align: 'Center', vertical: 'Middle', margins: { left: PAD, right: PAD, top: 0, bottom: 0 },
      }),
    })),
  }
}

/**
 * A two-by-two matrix over the left six columns: four quadrants, the top right in navy, and axis labels. The vertical
 * label is turned to read upwards when PowerPoint can rotate shapes; otherwise it sits above the matrix.
 */
export function matrixSpec(rotate: boolean): ElementSpec {
  const area = columns(0, 5)
  const left = area.left + LABEL + PAD
  const width = area.width - LABEL - PAD
  const top = CONTENT_AREA.top + (rotate ? 0 : LABEL + PAD)
  const height = CONTENT_AREA.height - LABEL - PAD - (rotate ? 0 : LABEL + PAD)
  const qw = (width - PAD) / 2
  const qh = (height - PAD) / 2
  const label = (t: string): TextSpec => text([{ text: t, font: small({ colour: COLOURS.charcoal }) }], { align: 'Center', vertical: 'Middle', margins: NO_MARGINS })
  const quadrants = [[0, 0], [1, 0], [0, 1], [1, 1]].map(([c = 0, r = 0], i) => {
    const navy = c === 1 && r === 0
    const at = { left: left + c * (qw + PAD), top: top + r * (qh + PAD), width: qw, height: qh }
    return rect(`BwB matrix quadrant ${i + 1}`, at, navy ? COLOURS.navy : COLOURS.offWhite, {
      text: text([
        { text: 'Quadrant title', font: body({ bold: true, colour: navy ? COLOURS.white : COLOURS.navy }) },
        { text: 'What belongs here', font: small({ colour: navy ? COLOURS.white : COLOURS.navy }) },
      ], { margins: { left: PAD, right: PAD, top: PAD, bottom: PAD } }),
    })
  })
  const xLabel: ShapeSpec = {
    kind: 'text', geometry: 'Rectangle', name: 'BwB matrix horizontal axis', left, top: top + height + PAD, width, height: LABEL,
    fill: null, line: null, text: label('Horizontal axis: low to high →'),
  }
  const yText = rotate ? 'Vertical axis: low to high →' : '↑ Vertical axis: low to high'
  const yLabel: ShapeSpec = rotate
    ? {
      ...xLabel, name: 'BwB matrix vertical axis', left: area.left + LABEL / 2 - height / 2, top: top + height / 2 - LABEL / 2, width: height,
      rotation: 270, text: label(yText),
    }
    : { ...xLabel, name: 'BwB matrix vertical axis', top: CONTENT_AREA.top, text: { ...label(yText), align: 'Left' } }
  return { kind: 'matrix', name: 'BwB matrix', group: true, shapes: [...quadrants, xLabel, yLabel] }
}

/** A navy line across the content area with four milestones, each with a label and a date below. */
export function timelineSpec(): ElementSpec {
  const y = CONTENT_AREA.top + GUTTER
  const dot = GUTTER
  const step = CONTENT_AREA.width / 4
  const line: ShapeSpec = {
    kind: 'line', geometry: 'Rectangle', name: 'BwB timeline line', left: CONTENT_AREA.left, top: y, width: CONTENT_AREA.width, height: 0,
    fill: null, line: { colour: COLOURS.navy, weight: AXIS_WEIGHT },
  }
  const milestones = [0, 1, 2, 3].flatMap((i): ShapeSpec[] => {
    const x = CONTENT_AREA.left + step * (i + 0.5)
    return [
      { ...rect(`BwB milestone ${i + 1}`, { left: x - dot / 2, top: y - dot / 2, width: dot, height: dot }, COLOURS.navy), geometry: 'Ellipse' },
      {
        kind: 'text', geometry: 'Rectangle', name: `BwB milestone ${i + 1} label`, left: x - (step - GUTTER) / 2, top: y + dot / 2 + PAD,
        width: step - GUTTER, height: SLIDE_TYPE.body.size * LINE + SLIDE_TYPE.small.size * LINE + PAD, fill: null, line: null,
        text: text([{ text: 'Milestone', font: body({ bold: true }) }, { text: 'Date', font: small({ colour: COLOURS.charcoal }) }], {
          align: 'Center', margins: NO_MARGINS,
        }),
      },
    ]
  })
  return { kind: 'timeline', name: 'BwB timeline', group: true, shapes: [line, ...milestones] }
}

/** Navy circles numbered 1 to 4, marker sized, in a row at the top left of the content area. Each moves on its own. */
export function stepMarkersSpec(): ElementSpec {
  const size = cmToPt(ICON_SIZES_CM.marker)
  return {
    kind: 'step-marker', name: 'BwB step markers', group: false,
    shapes: [1, 2, 3, 4].map((n, i) => ({
      ...rect(`BwB step marker ${n}`, { left: CONTENT_AREA.left + i * (size + PAD), top: CONTENT_AREA.top, width: size, height: size }, COLOURS.navy),
      geometry: 'Ellipse',
      text: text([{ text: String(n), font: small({ colour: COLOURS.white, bold: true }) }], { align: 'Center', vertical: 'Middle', margins: NO_MARGINS, wrap: false }),
    })),
  }
}

/** The source line: Aptos 8 pt charcoal in the source area under the content. */
export function sourceLineSpec(): ElementSpec {
  return {
    kind: 'source', name: 'BwB source line', group: false,
    shapes: [{
      kind: 'text', geometry: 'Rectangle', name: 'BwB source line', ...SOURCE_AREA, fill: null, line: null,
      text: text([{
        text: 'Source: organisation, title (year)',
        font: { name: SLIDE_TYPE.source.font, size: SLIDE_TYPE.source.size, colour: SLIDE_TYPE.source.colour },
      }], {
        vertical: 'Bottom', margins: NO_MARGINS,
      }),
    }],
  }
}

/** A navy review note over the right third of the title, nudged down and left past notes already there. */
export function noteSpec(date: string, existing = 0): ElementSpec {
  const area = columns(8, 11)
  const shift = existing * PAD
  return {
    kind: 'note', name: 'BwB review note', group: false,
    shapes: [rect('BwB review note', {
      left: area.left - shift, top: MARGINS.top + shift, width: area.width, height: 3 * SLIDE_TYPE.small.size * LINE + PAD,
    }, COLOURS.navy, {
      text: text([{ text: 'Note: ', font: small({ colour: COLOURS.white, bold: true }) }, { text: date, font: small({ colour: COLOURS.white }) }], {
        margins: { left: PAD, right: PAD, top: PAD / 2, bottom: PAD / 2 },
      }),
    })],
  }
}

/** Upper-case Aptos bold runs about 0.7 of the type size per letter. */
const CAPS_WIDTH = 0.7

/** A status sticker such as DRAFT at the top right, in the band above the title. */
export function stickerSpec(label: string, slideWidth: number = SLIDE.width): ShapeSpec {
  const words = label.toUpperCase()
  const size = SLIDE_TYPE.small.size
  const height = size * LINE + PAD / 2
  const width = words.length * size * CAPS_WIDTH + 2 * PAD
  return rect('BwB sticker', { left: slideWidth - MARGINS.right - width, top: (MARGINS.top - height) / 2, width, height }, COLOURS.white, {
    line: { colour: COLOURS.charcoal, weight: RULE.weight },
    text: text([{ text: words, font: small({ colour: COLOURS.charcoal, bold: true }) }], {
      align: 'Center', vertical: 'Middle', margins: { left: PAD, right: PAD, top: 0, bottom: 0 }, wrap: false,
    }),
  })
}

/** Red, amber and green for trackers. Amber and red are outside the BwB palette, so these dots carry BWB_RAG. */
export const RAG_COLOURS = { green: COLOURS.green, amber: 'F2A900', red: 'A72C28' } as const
export type RagColour = keyof typeof RAG_COLOURS

export function ragSpec(colour: RagColour, at: { left: number; top: number }): ShapeSpec {
  const size = cmToPt(ICON_SIZES_CM.inline)
  return { ...rect(`BwB RAG ${colour}`, { ...at, width: size, height: size }, RAG_COLOURS[colour]), geometry: 'Ellipse' }
}

/**
 * Where the next mark (Harvey ball, status mark or RAG dot) goes: beside the mark selected now, so clicking a few
 * times builds a row (wrapping below at the right margin), or else in the middle of the content area.
 */
export function markPosition(previous: Rect | null, size: number, slideWidth: number = SLIDE.width): { left: number; top: number } {
  if (!previous) return { left: CONTENT_AREA.left + (CONTENT_AREA.width - size) / 2, top: CONTENT_AREA.top + (CONTENT_AREA.height - size) / 2 }
  const left = previous.left + previous.width + PAD
  if (left + size <= slideWidth - MARGINS.right) return { left, top: previous.top + (previous.height - size) / 2 }
  return { left: previous.left, top: previous.top + previous.height + PAD }
}

export interface TableSpec {
  rows: number
  columns: number
  values: string[][]
  left: number
  top: number
  width: number
  header: { fill: string; font: FontSpec }
  body: { fill: string; font: FontSpec }
  rule: { colour: string; weight: number }
}

/** A four-by-four table across the content area: a navy header row in white bold, light grey rules. */
export function tableSpec(): TableSpec {
  const rows = 4
  const cols = 4
  const font = small({ colour: COLOURS.black })
  return {
    rows, columns: cols,
    values: Array.from({ length: rows }, (_, r) => Array.from({ length: cols }, () => (r === 0 ? 'Heading' : ''))),
    left: CONTENT_AREA.left, top: CONTENT_AREA.top, width: CONTENT_AREA.width,
    header: { fill: COLOURS.navy, font: { ...font, colour: COLOURS.white, bold: true } },
    body: { fill: COLOURS.white, font },
    rule: RULE,
  }
}
