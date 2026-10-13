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

/**
 * Data map colours, light to dark: tints of BwB green (20%, 40%, 50%, 60% and 80% on white), the green itself, and a
 * shade halfway from green to navy for the darkest class. Countries without data are light grey; neighbours outside
 * the region are off-white.
 */
export const MAP_TINTS = {
  green20: 'D6E7D9',
  green40: 'ADCFB3',
  green50: '99C4A0',
  green60: '84B88C',
  green80: '5BA066',
  green: COLOURS.green,
  greenNavy: '1A563C',
} as const

/** Classes for 3, 4 and 5 bands, light to dark. */
export const MAP_RAMPS: Record<3 | 4 | 5, readonly string[]> = {
  3: [MAP_TINTS.green20, MAP_TINTS.green60, MAP_TINTS.green],
  4: [MAP_TINTS.green20, MAP_TINTS.green50, MAP_TINTS.green, MAP_TINTS.greenNavy],
  5: [MAP_TINTS.green20, MAP_TINTS.green40, MAP_TINTS.green60, MAP_TINTS.green, MAP_TINTS.greenNavy],
}

/** Every colour a chart may contain: the series colours and white. */
export const CHART_HEX: ReadonlySet<string> = new Set([...SERIES_HEX, COLOURS.white])

/** Every colour a data map may contain: the chart colours, the green tints and off-white. */
export const MAP_HEX: ReadonlySet<string> = new Set([...CHART_HEX, COLOURS.offWhite, ...Object.values(MAP_TINTS)])

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
