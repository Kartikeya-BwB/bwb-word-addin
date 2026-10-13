// SVG for the marks the add-in draws as graphics: Harvey balls and the Lucide status marks. PowerPoint inserts SVG as
// a vector graphic people can recolour or convert to shapes.

import type { IconColourName, IconIndex } from '../brand/icons.ts'
import { iconSvg } from '../brand/icons.ts'
import { COLOURS } from '../brand/tokens.ts'

export type Quarters = 0 | 1 | 2 | 3 | 4

const ARC_ENDS: Record<1 | 2 | 3, string> = { 1: 'A9 9 0 0 1 21 12', 2: 'A9 9 0 0 1 12 21', 3: 'A9 9 0 1 1 3 12' }

/** A Harvey ball: a navy ring filled clockwise from the top by quarters, as on the ribbon (scripts/ribbon-lib.mjs). */
export function harveySvg(quarters: Quarters, hex: string = COLOURS.navy): string {
  const colour = `#${hex}`
  const fill = quarters === 4
    ? `<circle cx="12" cy="12" r="9" fill="${colour}"/>`
    : quarters === 0 ? '' : `<path d="M12 12 L12 3 ${ARC_ENDS[quarters]} Z" fill="${colour}"/>`
  return `<svg xmlns="http://www.w3.org/2000/svg" width="24" height="24" viewBox="0 0 24 24">${fill}` +
    `<circle cx="12" cy="12" r="9" fill="none" stroke="${colour}" stroke-width="2"/></svg>`
}

export const HARVEY_LABELS: Record<Quarters, string> = {
  0: 'Harvey ball, empty (0 per cent)',
  1: 'Harvey ball, a quarter (25 per cent)',
  2: 'Harvey ball, half (50 per cent)',
  3: 'Harvey ball, three quarters (75 per cent)',
  4: 'Harvey ball, full (100 per cent)',
}

export type StatusKind = 'onTrack' | 'atRisk' | 'offTrack'

export const STATUS_MARKS: Record<StatusKind, { icon: string; colour: IconColourName; label: string }> = {
  onTrack: { icon: 'circle-check', colour: 'green', label: 'On track' },
  atRisk: { icon: 'circle-alert', colour: 'charcoal', label: 'At risk' },
  offTrack: { icon: 'circle-x', colour: 'charcoal', label: 'Off track' },
}

/** The Lucide status mark in its BwB colour, from the bundled icon index. */
export function statusSvg(index: IconIndex, kind: StatusKind): string {
  const mark = STATUS_MARKS[kind]
  const svg = iconSvg(index, mark.icon, mark.colour)
  if (!svg) throw new Error(`The icon ${mark.icon} is missing from the icon library.`)
  return svg
}
