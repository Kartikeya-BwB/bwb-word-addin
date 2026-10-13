// Chart colours from the BwB palette (src/brand/tokens.ts). Series take navy, green, green 50, charcoal and light
// grey in that order; charcoal 20 is a sixth colour reached only by a sixth series or donut segment. Lime belongs on
// navy backgrounds only, so it never appears in a chart drawn on white.

import { COLOURS } from '../brand/tokens.ts'

export const SERIES_HEX: readonly string[] = [
  COLOURS.navy,
  COLOURS.green,
  COLOURS.green50,
  COLOURS.charcoal,
  COLOURS.lightGrey,
  COLOURS.charcoal20,
]

export const PALETTE_NAMES: readonly string[] = ['Navy', 'Green', 'Green 50', 'Charcoal', 'Light grey', 'Charcoal 20']

/** Ink for the parts of a chart that are not data. */
export const INK = {
  /** Axis labels, category labels, legend text, the source line and the zero line. */
  text: COLOURS.charcoal,
  /** Data labels. */
  data: COLOURS.navy,
  grid: COLOURS.lightGrey,
  white: COLOURS.white,
  navy: COLOURS.navy,
  green: COLOURS.green,
  charcoal: COLOURS.charcoal,
} as const

/** Every colour a chart may contain: the series colours and white. */
export const CHART_HEX: ReadonlySet<string> = new Set([...SERIES_HEX, COLOURS.white])

/** The palette colour at an index, wrapping round. */
export function paletteColour(index: number): string {
  const n = SERIES_HEX.length
  const i = ((Math.trunc(index) % n) + n) % n
  return SERIES_HEX[i] ?? COLOURS.navy
}

/** Text that reads on a fill: white on navy, green and charcoal, navy on the light tints. */
export function inkOn(fill: string): string {
  return fill === COLOURS.navy || fill === COLOURS.green || fill === COLOURS.charcoal ? COLOURS.white : COLOURS.navy
}
