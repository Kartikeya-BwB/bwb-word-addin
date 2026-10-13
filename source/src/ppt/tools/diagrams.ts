// Diagrams as native shapes in BwB colours with text to type over: pyramid, funnel, cycle, Venn, and hub and spoke.
// Plain layouts in points, centred in the content area; tools/draw.ts puts them on the slide.

import { CONTENT_AREA, SLIDE_TYPE } from '../../brand/slideGrid.ts'
import { COLOURS } from '../../brand/tokens.ts'
import type { Rect } from '../../geometry/index.ts'
import { UserFacingError } from '../../word/run.ts'
import type { TextSpec } from '../specs.ts'
import type { DrawShape } from './draw.ts'
import { insertDrawing, segment } from './draw.ts'

export type DiagramKind = 'pyramid' | 'funnel' | 'cycle' | 'venn' | 'hub-spoke'

export interface DiagramInfo {
  kind: DiagramKind
  label: string
  hint: string
  min: number
  max: number
  initial: number
  /** What the items are called: levels, stages, steps, circles, spokes. */
  noun: string
}

export const DIAGRAMS: DiagramInfo[] = [
  { kind: 'pyramid', label: 'Pyramid', hint: 'Levels from the top down', min: 3, max: 5, initial: 4, noun: 'levels' },
  { kind: 'funnel', label: 'Funnel', hint: 'Stages that narrow towards the end', min: 3, max: 5, initial: 4, noun: 'stages' },
  { kind: 'cycle', label: 'Cycle', hint: 'Steps around a circle', min: 3, max: 6, initial: 4, noun: 'steps' },
  { kind: 'venn', label: 'Venn', hint: 'Overlapping circles', min: 2, max: 3, initial: 3, noun: 'circles' },
  { kind: 'hub-spoke', label: 'Hub and spoke', hint: 'A centre with spokes around it', min: 4, max: 8, initial: 6, noun: 'spokes' },
]

/** Fills from the top level down, with the text colour each takes. Lime stays off: it belongs on navy only. */
const SHADES = [
  { fill: COLOURS.navy, text: COLOURS.white },
  { fill: COLOURS.green, text: COLOURS.white },
  { fill: COLOURS.green50, text: COLOURS.navy },
  { fill: COLOURS.lightGrey, text: COLOURS.navy },
  { fill: COLOURS.offWhite, text: COLOURS.navy },
]
const shade = (i: number): { fill: string; text: string } => SHADES[i % SHADES.length] ?? { fill: COLOURS.navy, text: COLOURS.white }

const APTOS = SLIDE_TYPE.body.font
const NO_MARGINS = { left: 0, right: 0, top: 0, bottom: 0 }
/** The white rules between levels, in points. */
const SEPARATOR = 3

function label(text: string, colour: string, size: number = SLIDE_TYPE.small.size): TextSpec {
  return { paragraphs: [{ text, font: { name: APTOS, size, colour, bold: true } }], align: 'Center', vertical: 'Middle', margins: NO_MARGINS, wrap: true }
}

const textBox = (name: string, cx: number, cy: number, width: number, height: number, text: TextSpec): DrawShape => ({
  kind: 'text', name, left: cx - width / 2, top: cy - height / 2, width, height, fill: null, line: null, text,
})

/**
 * Levels as stacked triangles that share the apex, largest at the back: each shows as a band with the pyramid's own
 * slope, which needs no shape adjustments. White rules part the levels; each level's text sits in its band.
 */
function pyramid(n: number, area: Rect): DrawShape[] {
  const height = area.height
  const width = Math.min(area.width, height * 1.6)
  const cx = area.left + area.width / 2
  const top = area.top
  const out: DrawShape[] = []
  for (let k = n; k >= 1; k -= 1) {
    out.push({
      kind: 'geometric', geometry: 'Triangle', name: `BwB pyramid level ${k}`, left: cx - (width * k) / (2 * n), top,
      width: (width * k) / n, height: (height * k) / n, fill: { colour: shade(k - 1).fill }, line: null,
    })
  }
  for (let k = 1; k < n; k += 1) {
    const y = top + (height * k) / n
    out.push(segment(`BwB pyramid rule ${k}`, cx - (width * k) / (2 * n), y, cx + (width * k) / (2 * n), y, { colour: COLOURS.white, weight: SEPARATOR }))
  }
  for (let k = 1; k <= n; k += 1) {
    const across = (width * (k - 0.4)) / n
    out.push(textBox(`BwB pyramid level ${k} text`, cx, top + (height * (k - 0.4)) / n, across * 0.8, (height / n) * 0.6, label(`Level ${k}`, shade(k - 1).text)))
  }
  return out
}

/** Stages as stacked downward triangles that share their tip: an upside-down pyramid, the first stage widest. */
function funnel(n: number, area: Rect): DrawShape[] {
  const height = area.height
  const width = Math.min(area.width, height * 1.6)
  const cx = area.left + area.width / 2
  const top = area.top
  const out: DrawShape[] = []
  for (let j = 0; j < n; j += 1) {
    out.push({
      kind: 'geometric', geometry: 'FlowChartMerge', name: `BwB funnel stage ${j + 1}`, left: cx - (width * (n - j)) / (2 * n),
      top: top + (height * j) / n, width: (width * (n - j)) / n, height: (height * (n - j)) / n, fill: { colour: shade(j).fill }, line: null,
    })
  }
  for (let j = 1; j < n; j += 1) {
    const y = top + (height * j) / n
    out.push(segment(`BwB funnel rule ${j}`, cx - (width * (n - j)) / (2 * n), y, cx + (width * (n - j)) / (2 * n), y, { colour: COLOURS.white, weight: SEPARATOR }))
  }
  for (let j = 0; j < n; j += 1) {
    const across = (width * (n - j - 0.4)) / n
    out.push(textBox(`BwB funnel stage ${j + 1} text`, cx, top + (height * (j + 0.4)) / n, across * 0.8, (height / n) * 0.6, label(`Stage ${j + 1}`, shade(j).text)))
  }
  return out
}

const degrees = (radians: number): number => ((((radians * 180) / Math.PI) % 360) + 360) % 360

/** Navy circles on a light grey ring, first at the top, with green arrows between them turned along the ring. */
function cycle(n: number, area: Rect): DrawShape[] {
  const cx = area.left + area.width / 2
  const cy = area.top + area.height / 2
  const r = Math.min(area.width, area.height) * 0.34
  const chord = 2 * r * Math.sin(Math.PI / n)
  const d = Math.min(96, chord - 30)
  const out: DrawShape[] = [
    { kind: 'geometric', geometry: 'Ellipse', name: 'BwB cycle ring', left: cx - r, top: cy - r, width: 2 * r, height: 2 * r, fill: null, line: { colour: COLOURS.lightGrey, weight: 2 } },
  ]
  for (let k = 0; k < n; k += 1) {
    const phi = -Math.PI / 2 + (2 * Math.PI * (k + 0.5)) / n
    out.push({
      kind: 'geometric', geometry: 'RightArrow', name: `BwB cycle arrow ${k + 1}`, left: cx + r * Math.cos(phi) - 11, top: cy + r * Math.sin(phi) - 8,
      width: 22, height: 16, rotation: degrees(phi + Math.PI / 2), requiresRotation: true, fill: { colour: COLOURS.green }, line: null,
    })
  }
  for (let k = 0; k < n; k += 1) {
    const theta = -Math.PI / 2 + (2 * Math.PI * k) / n
    out.push({
      kind: 'geometric', geometry: 'Ellipse', name: `BwB cycle step ${k + 1}`, left: cx + r * Math.cos(theta) - d / 2, top: cy + r * Math.sin(theta) - d / 2,
      width: d, height: d, fill: { colour: COLOURS.navy }, line: null, text: label(`Step ${k + 1}`, COLOURS.white),
    })
  }
  return out
}

/** Two or three overlapping circles in light tints with full-colour outlines, so overlaps read darker. */
function venn(n: number, area: Rect): DrawShape[] {
  const d = n === 2 ? 240 : 220
  const cx = area.left + area.width / 2
  const cy = area.top + area.height / 2
  const centres = n === 2
    ? [[cx - 0.32 * d, cy], [cx + 0.32 * d, cy]]
    : [[cx - 0.32 * d, cy - 0.27 * d], [cx + 0.32 * d, cy - 0.27 * d], [cx, cy + 0.27 * d]]
  const colours = [COLOURS.navy, COLOURS.green, COLOURS.charcoal]
  const mx = centres.reduce((s, [x = 0]) => s + x, 0) / centres.length
  const my = centres.reduce((s, [, y = 0]) => s + y, 0) / centres.length
  const out: DrawShape[] = centres.map(([x = 0, y = 0], i) => ({
    kind: 'geometric', geometry: 'Ellipse', name: `BwB Venn circle ${i + 1}`, left: x - d / 2, top: y - d / 2, width: d, height: d,
    fill: { colour: colours[i] ?? COLOURS.navy, transparency: 0.6 }, line: { colour: colours[i] ?? COLOURS.navy, weight: 1 },
  }))
  centres.forEach(([x = 0, y = 0], i) => {
    const len = Math.hypot(x - mx, y - my) || 1
    const lx = x + ((x - mx) / len) * 0.22 * d
    const ly = y + ((y - my) / len) * 0.22 * d
    out.push(textBox(`BwB Venn label ${i + 1}`, lx, ly, 0.45 * d, 0.2 * d, label(`Set ${String.fromCharCode(65 + i)}`, COLOURS.navy, SLIDE_TYPE.body.size)))
  })
  out.push(textBox('BwB Venn overlap', mx, my, 0.3 * d, 0.12 * d, label('Shared', COLOURS.navy)))
  return out
}

/** A navy hub with spokes around it on an ellipse, each joined to the hub by a line. */
function hubAndSpoke(n: number, area: Rect): DrawShape[] {
  const cx = area.left + area.width / 2
  const cy = area.top + area.height / 2
  const hub = 112
  const w = 136
  const h = 48
  const rx = area.width / 2 - w / 2 - 8
  const ry = area.height / 2 - h / 2 - 8
  const lines: DrawShape[] = []
  const spokes: DrawShape[] = []
  for (let k = 0; k < n; k += 1) {
    const theta = -Math.PI / 2 + (2 * Math.PI * k) / n
    const sx = cx + rx * Math.cos(theta)
    const sy = cy + ry * Math.sin(theta)
    const len = Math.hypot(sx - cx, sy - cy) || 1
    const ux = (sx - cx) / len
    const uy = (sy - cy) / len
    const reach = Math.min(Math.abs(ux) > 1e-9 ? w / 2 / Math.abs(ux) : Infinity, Math.abs(uy) > 1e-9 ? h / 2 / Math.abs(uy) : Infinity)
    lines.push(segment(`BwB spoke line ${k + 1}`, cx + (ux * hub) / 2, cy + (uy * hub) / 2, sx - ux * reach, sy - uy * reach, { colour: COLOURS.charcoal, weight: 1 }))
    spokes.push({
      kind: 'geometric', geometry: 'RoundRectangle', name: `BwB spoke ${k + 1}`, left: sx - w / 2, top: sy - h / 2, width: w, height: h,
      fill: { colour: COLOURS.white }, line: { colour: COLOURS.navy, weight: 1 }, text: label(`Spoke ${k + 1}`, COLOURS.navy),
    })
  }
  const centre: DrawShape = {
    kind: 'geometric', geometry: 'Ellipse', name: 'BwB hub', left: cx - hub / 2, top: cy - hub / 2, width: hub, height: hub,
    fill: { colour: COLOURS.navy }, line: null, text: label('Hub', COLOURS.white, SLIDE_TYPE.body.size),
  }
  return [...lines, centre, ...spokes]
}

/** The shapes of a diagram with `count` items (clamped to what the diagram allows), in the content area by default. */
export function diagramShapes(kind: DiagramKind, count: number, area: Rect = CONTENT_AREA): DrawShape[] {
  const info = DIAGRAMS.find((d) => d.kind === kind)
  const n = Math.round(Math.min(info?.max ?? count, Math.max(info?.min ?? count, count)))
  switch (kind) {
    case 'pyramid': return pyramid(n, area)
    case 'funnel': return funnel(n, area)
    case 'cycle': return cycle(n, area)
    case 'venn': return venn(n, area)
    case 'hub-spoke': return hubAndSpoke(n, area)
  }
}

/** Puts a diagram on the current slide, grouped and tagged BWB_ELEMENT with its kind. */
export async function insertDiagram(kind: DiagramKind, count: number): Promise<string> {
  const info = DIAGRAMS.find((d) => d.kind === kind)
  if (!info) throw new Error(`Unknown diagram ${kind}`)
  if (!Number.isInteger(count) || count < info.min || count > info.max) {
    throw new UserFacingError(`A ${info.label.toLowerCase()} takes ${info.min} to ${info.max} ${info.noun}.`)
  }
  const drawn = await insertDrawing(kind, `BwB ${info.label.toLowerCase()}`, diagramShapes(kind, count), `The ${info.label.toLowerCase()} diagram`)
  const skipped = drawn.skipped ? ' This PowerPoint cannot turn shapes, so the arrows and slanted lines were left out.' : ''
  return `${info.label} with ${count} ${info.noun} added. Type over the text.${skipped}`
}
