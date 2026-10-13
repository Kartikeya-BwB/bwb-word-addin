/**
 * The BwB icon rules: approved colours, the substitutions the builder makes (and the warnings it gives), Lucide name
 * matching with older names, and the {icon:name:colour} token. A port of the rules in bwb_lib/icons.py.
 *
 * Approved treatments: green, navy and charcoal on white or off-white; white on navy (the default there); lime on navy
 * only, one accent per panel.
 */
import { LUCIDE_ALIASES } from './lucideAliases.ts'
import { getCloseMatches } from './similar.ts'
import type { IconColour } from './types.ts'
import { esc, pyRound, pyStrip } from './xml.ts'

export const ICON_COLOURS: Readonly<Record<IconColour, string>> = {
  navy: '022338',
  green: '328840',
  lime: 'BED62F',
  charcoal: '4A4A4C',
  white: 'FFFFFF',
}
const LIGHT_OK: readonly IconColour[] = ['green', 'navy', 'charcoal']
const DARK_OK: readonly IconColour[] = ['white', 'lime']
const COLOUR_ALIASES = new Map([
  ['grey', 'charcoal'],
  ['gray', 'charcoal'],
  ['black', 'navy'],
  ['primary', 'green'],
  ['dark', 'navy'],
])

/** Brand sizes (cm): inline with body text, callout header, standalone marker, large panel icon. */
export const ICON_SIZES_CM = { inline: 0.4, callout: 0.6, marker: 0.8, panel: 1.2 } as const

/** Width in pixels of the PNG fallback the builder embeds beside each SVG icon. */
export const ICON_PX = 256

function isIconColour(c: string): c is IconColour {
  return Object.hasOwn(ICON_COLOURS, c)
}

/** The approved colour a name, alias (grey, black ...) or hex value stands for, or null. */
export function colourName(colour: string | null | undefined): IconColour | null {
  const raw = pyStrip(colour ?? '').toLowerCase()
  const c = COLOUR_ALIASES.get(raw) ?? raw
  if (isIconColour(c)) return c
  const hex = c.replace(/^#+/, '')
  for (const k of Object.keys(ICON_COLOURS) as IconColour[]) if (ICON_COLOURS[k].toLowerCase() === hex) return k
  return null
}

export interface ColourCheck {
  colour: IconColour
  /** Why the requested colour was replaced, or null when it was allowed. */
  note: string | null
}

/**
 * The colour to use for an icon on a navy panel or a light background, with a note when the request breaks the brand
 * rules. `fallback` replaces a disallowed colour (default white on navy, green on light).
 */
export function checkColour(colour: string | null | undefined, onNavy: boolean, fallback?: string | null): ColourCheck {
  const f = colourName(fallback)
  const fb: IconColour = f && DARK_OK.includes(f) === onNavy ? f : onNavy ? 'white' : 'green'
  const c = colourName(colour)
  if (c === null) return { colour: fb, note: `'${String(colour)}' is not an approved icon colour; used ${fb}` }
  if (onNavy && LIGHT_OK.includes(c)) {
    return {
      colour: 'white',
      note:
        c !== 'charcoal'
          ? `${c} icon on a navy panel; used white (white or lime only on navy)`
          : 'charcoal icon on a navy panel is hard to read; used white',
    }
  }
  if (!onNavy && DARK_OK.includes(c)) {
    return { colour: fb, note: `${c} icon on a light background; used ${fb} (lime and white are for navy panels only)` }
  }
  return { colour: c, note: null }
}

/** "Lucide_Leaf Icon" becomes "leaf": lower case, hyphens, no "lucide-" prefix or "-icon" suffix. */
export function normaliseIconName(name: string | null | undefined): string {
  let s = pyStrip(String(name ?? ''))
    .toLowerCase()
    .replaceAll('_', '-')
    .replaceAll(' ', '-')
  s = s.replace(/^lucide[-:]/, '')
  s = s.replace(/-icon$/, '')
  return s
}

/**
 * The current Lucide name for `name`, following renames (bar-chart becomes chart-no-axes-column-increasing), or null.
 * `isIcon` says whether a normalised name is a current icon; `aliases` maps older names to current ones.
 */
export function resolveIconName(
  name: string,
  isIcon: (name: string) => boolean,
  aliases: Readonly<Record<string, string>> = LUCIDE_ALIASES,
): string | null {
  const s = normaliseIconName(name)
  if (isIcon(s)) return s
  return Object.hasOwn(aliases, s) ? (aliases[s] ?? null) : null
}

/** A resolve function for an IconRenderer, from the current icon names (and optionally another alias table). */
export function createIconResolver(
  names: Iterable<string>,
  aliases: Readonly<Record<string, string>> = LUCIDE_ALIASES,
): (name: string) => string | null {
  const set = new Set(names)
  return (name) => resolveIconName(name, (n) => set.has(n), aliases)
}

/**
 * Up to n suggestions for an unknown name: close matches (difflib, cutoff 0.6) then names containing it, each mapped to
 * its current name. `pool` is every icon name followed by the older names, in index order.
 */
export function suggestIconNames(
  name: string,
  pool: readonly string[],
  resolve: (name: string) => string | null,
  n = 5,
): string[] {
  const s = normaliseIconName(name)
  const close = getCloseMatches(s, pool, n, 0.6)
  const sub = s ? pool.filter((p) => p.includes(s)).slice(0, n) : []
  const seen = new Set<string>()
  const out: string[] = []
  for (const c of [...close, ...sub]) {
    const r = resolve(c) ?? c
    if (!seen.has(r)) {
      seen.add(r)
      out.push(r)
    }
  }
  return out.slice(0, n)
}

/** A Lucide icon as data: [element, attributes] pairs (lucide-static's icon-nodes.json, the add-in's lucide index). */
export type LucideIconNode = ReadonlyArray<readonly [string, Readonly<Record<string, string>>]>

const LUCIDE_SVG_OPEN =
  '<svg xmlns="http://www.w3.org/2000/svg" width="24" height="24" viewBox="0 0 24 24" fill="none" ' +
  'stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round">'

/** The standard Lucide SVG for an icon node, ready for recolourSvg. */
export function lucideSvgFromNodes(nodes: LucideIconNode): string {
  const shapes = nodes.map(([tag, a]) => `<${tag} ${Object.entries(a).map(([k, v]) => `${k}="${esc(v)}"`).join(' ')}/>`)
  return LUCIDE_SVG_OPEN + shapes.join('') + '</svg>'
}

/** Lucide SVG text in a BwB colour (the stroke is the only colour in a Lucide icon). */
export function recolourSvg(svg: string, colour: IconColour, strokeWidth?: number): string {
  let text = svg.replaceAll('stroke="currentColor"', `stroke="#${ICON_COLOURS[colour]}"`)
  if (strokeWidth) text = text.replaceAll('stroke-width="2"', `stroke-width="${strokeWidth}"`)
  return text
}

/** Pixels for a 600 dpi rendering of an icon `cm` wide (never below 96). */
export function pxForCm(cm: number): number {
  return Math.max(96, pyRound((cm / 2.54) * 600))
}

// ---- inline mark-up: {icon:name} or {icon:name:colour} ----

export const ICON_TOKEN = '\\{icon:[A-Za-z0-9][A-Za-z0-9_\\- ]*(?::[A-Za-z0-9#]+)?\\}'
const ICON_FULL = new RegExp(`^${ICON_TOKEN}$`)
const ICON_ALL = new RegExp(ICON_TOKEN, 'g')

export function isIconToken(s: string): boolean {
  return ICON_FULL.test(s)
}

/** '{icon:leaf:navy}' gives { name: 'leaf', colour: 'navy' }; '{icon:leaf}' gives colour null. */
export function parseIconToken(tok: string): { name: string; colour: string | null } {
  const body = tok.slice('{icon:'.length, -1)
  const i = body.indexOf(':')
  const name = i < 0 ? body : body.slice(0, i)
  const colour = i < 0 ? '' : body.slice(i + 1)
  return { name: pyStrip(name), colour: pyStrip(colour) || null }
}

export function stripIconTokens(text: string): string {
  return String(text).replace(ICON_ALL, '')
}
