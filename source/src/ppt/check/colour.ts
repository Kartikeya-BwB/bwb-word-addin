// The BwB palette for slides: which colours are allowed, which may colour text, and the nearest palette colour to an
// off-brand one, measured in CIELAB so a stray blue maps to navy rather than to whichever grey is closest in RGB.

import type { ColourName } from '../../brand/tokens.ts'
import { ALLOWED_HEX, COLOURS, PALETTE } from '../../brand/tokens.ts'

export { ALLOWED_HEX }

/** Colours text may take on a slide: navy, charcoal and black on light backgrounds, white on navy, green when large. */
export const TEXT_COLOURS: ColourName[] = ['navy', 'charcoal', 'black', 'white', 'green']

/** Every palette colour, for fills and outlines. */
export const SHAPE_COLOURS: ColourName[] = PALETTE.map((p) => p.key)

/** Backgrounds lime text fails on. */
const LIGHT_BACKGROUNDS = new Set<string>([COLOURS.white, COLOURS.offWhite, COLOURS.lightGrey, COLOURS.charcoal20, COLOURS.green50, COLOURS.lime])

export function isPaletteColour(hex: string): boolean {
  return ALLOWED_HEX.has(hex.toUpperCase())
}

export function isLightBackground(hex: string): boolean {
  return LIGHT_BACKGROUNDS.has(hex.toUpperCase())
}

/** 'Navy' for 022338; '#123456' for a colour outside the palette. */
export function colourLabel(hex: string): string {
  const upper = hex.toUpperCase()
  const entry = PALETTE.find((p) => p.hex === upper)
  if (entry) return entry.name
  return upper === COLOURS.charcoal20 ? 'Charcoal 20' : `#${upper}`
}

type Lab = [number, number, number]

function channel(c: number): number {
  const v = c / 255
  return v <= 0.04045 ? v / 12.92 : ((v + 0.055) / 1.055) ** 2.4
}

function lab(hex: string): Lab {
  const n = Number.parseInt(hex, 16)
  const [r, g, b] = [channel((n >> 16) & 255), channel((n >> 8) & 255), channel(n & 255)]
  // sRGB to XYZ (D65), scaled to the white point.
  const x = (0.4124 * r + 0.3576 * g + 0.1805 * b) / 0.95047
  const y = 0.2126 * r + 0.7152 * g + 0.0722 * b
  const z = (0.0193 * r + 0.1192 * g + 0.9505 * b) / 1.08883
  const f = (t: number): number => (t > 216 / 24389 ? Math.cbrt(t) : (24389 / 27 * t + 16) / 116)
  return [116 * f(y) - 16, 500 * (f(x) - f(y)), 200 * (f(y) - f(z))]
}

/** CIE76 distance between two hex colours. */
export function colourDistance(a: string, b: string): number {
  const [l1, a1, b1] = lab(a)
  const [l2, a2, b2] = lab(b)
  return Math.hypot(l1 - l2, a1 - a2, b1 - b2)
}

/** The palette colour, from `among`, closest to `hex`. */
export function nearestColour(hex: string, among: ColourName[]): { key: ColourName; hex: string } {
  let best: { key: ColourName; hex: string; d: number } | null = null
  for (const key of among) {
    const d = colourDistance(hex, COLOURS[key])
    if (!best || d < best.d) best = { key, hex: COLOURS[key], d }
  }
  return best ? { key: best.key, hex: best.hex } : { key: 'navy', hex: COLOURS.navy }
}

/** WCAG 2 relative luminance of a hex colour. */
export function relativeLuminance(hex: string): number {
  const n = Number.parseInt(hex, 16)
  const linear = (c: number): number => {
    const v = c / 255
    return v <= 0.03928 ? v / 12.92 : ((v + 0.055) / 1.055) ** 2.4
  }
  return 0.2126 * linear((n >> 16) & 255) + 0.7152 * linear((n >> 8) & 255) + 0.0722 * linear(n & 255)
}

/** WCAG 2 contrast ratio between two hex colours, from 1 to 21. */
export function contrastRatio(a: string, b: string): number {
  const [hi, lo] = [relativeLuminance(a), relativeLuminance(b)].sort((x, y) => y - x)
  return ((hi ?? 0) + 0.05) / ((lo ?? 0) + 0.05)
}

/** WCAG AA: 4.5 to 1 for text, 3 to 1 for large text (18 pt, or 14 pt bold). */
export function contrastNeeded(size: number | null, bold: boolean | null): number {
  const large = size !== null && (size >= 18 || (size >= 14 && bold === true))
  return large ? 3 : 4.5
}
