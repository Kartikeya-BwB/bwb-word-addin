// The format painter's memory: the format picked up last, which parts of it to apply, and up to 12 named styles.
// All of it lives in localStorage (each read and write guarded), so a format picked up on one slide applies on any
// other, and the ribbon's commands, which run in a page of their own, see the same format as the pane.

import { PALETTE } from '../../brand/tokens.ts'
import { UserFacingError } from '../../word/run.ts'

export type FormatPart = 'position' | 'size' | 'rotation' | 'fillLine' | 'text' | 'frame' | 'adjustments'

export const FORMAT_PARTS: Array<{ id: FormatPart; label: string; version: string }> = [
  { id: 'position', label: 'Position', version: '1.4' },
  { id: 'size', label: 'Size', version: '1.4' },
  { id: 'rotation', label: 'Rotation', version: '1.10' },
  { id: 'fillLine', label: 'Fill and outline', version: '1.4' },
  { id: 'text', label: 'Font, size, colour, bold and italic', version: '1.4' },
  { id: 'frame', label: 'Text margins, autofit, anchor and wrap', version: '1.4' },
  { id: 'adjustments', label: 'Shape adjustments, such as corner rounding', version: '1.10' },
]

/** What a plain "apply" copies: the look, not the place or size. */
export const DEFAULT_PARTS: FormatPart[] = ['fillLine', 'text', 'frame']

export interface FontFormat {
  name: string | null
  size: number | null
  /** Hex without '#'. */
  colour: string | null
  bold: boolean | null
  italic: boolean | null
}

export interface FrameFormat {
  margins: { left: number; right: number; top: number; bottom: number }
  autoSize: string
  verticalAlignment: string
  wordWrap: boolean
}

export type FillFormat = { kind: 'solid'; colour: string; transparency: number } | { kind: 'none' }

export interface LineFormat {
  visible: boolean
  colour: string
  weight: number
  dashStyle: string
  transparency: number
}

export interface ShapeFormat {
  version: 1
  /** The name of the shape it was picked up from. */
  source: string
  /** Its type as PowerPoint reports it: GeometricShape, TextBox and so on. */
  type: string
  position: { left: number; top: number }
  size: { width: number; height: number }
  /** Null where PowerPoint does not report rotation (before PowerPointApi 1.10). */
  rotation: number | null
  /** Null for shapes without a fill, and for fills the API cannot set: gradients, patterns and pictures. */
  fill: FillFormat | null
  line: LineFormat | null
  /** The font of the first character: when the text is mixed, what comes first stands for the rest. */
  text: FontFormat | null
  frame: FrameFormat | null
  /** Null where PowerPoint does not report adjustments (before 1.10), or for shapes other than geometric shapes. */
  adjustments: number[] | null
}

export interface NamedStyle {
  id: string
  name: string
  format: ShapeFormat
}

export const MAX_STYLES = 12
const MAX_NAME = 40
const FORMAT_KEY = 'bwb.ppt.painter.format.v1'
const PARTS_KEY = 'bwb.ppt.painter.parts.v1'
const STYLES_KEY = 'bwb.ppt.painter.styles.v1'

function read(key: string): unknown {
  try {
    const raw = typeof localStorage === 'undefined' ? null : localStorage.getItem(key)
    return raw === null ? null : (JSON.parse(raw) as unknown)
  } catch {
    return null
  }
}

function write(key: string, value: unknown): void {
  try {
    if (typeof localStorage !== 'undefined') localStorage.setItem(key, JSON.stringify(value))
  } catch {
    // Storage refused (a private window, a full quota): the format lasts until the pane closes.
  }
}

let memory: ShapeFormat | null = null

export function isShapeFormat(value: unknown): value is ShapeFormat {
  const v = value as Partial<ShapeFormat> | null
  return Boolean(v) && v?.version === 1 && typeof v.source === 'string' && typeof v.position?.left === 'number' && typeof v.size?.width === 'number'
}

/** The format picked up last, from storage, or from this page's memory when storage is unavailable. */
export function loadFormat(): ShapeFormat | null {
  const stored = read(FORMAT_KEY)
  return isShapeFormat(stored) ? stored : memory
}

export function saveFormat(format: ShapeFormat): void {
  memory = format
  write(FORMAT_KEY, format)
}

export function loadParts(): FormatPart[] {
  const stored = read(PARTS_KEY)
  const known = new Set(FORMAT_PARTS.map((p) => p.id))
  return Array.isArray(stored) ? stored.filter((p): p is FormatPart => known.has(p as FormatPart)) : [...DEFAULT_PARTS]
}

export function saveParts(parts: readonly FormatPart[]): void {
  write(PARTS_KEY, [...parts])
}

export function loadStyles(): NamedStyle[] {
  const stored = read(STYLES_KEY)
  if (!Array.isArray(stored)) return []
  return stored.filter((s): s is NamedStyle => {
    const v = s as Partial<NamedStyle> | null
    return typeof v?.id === 'string' && typeof v.name === 'string' && isShapeFormat(v.format)
  }).slice(0, MAX_STYLES)
}

function cleanName(name: string, styles: readonly NamedStyle[], except?: string): string {
  const clean = name.replace(/\s+/g, ' ').trim()
  if (!clean) throw new UserFacingError('Give the style a name.')
  if (clean.length > MAX_NAME) throw new UserFacingError(`Keep style names to ${MAX_NAME} characters.`)
  if (styles.some((s) => s.id !== except && s.name.toLowerCase() === clean.toLowerCase())) {
    throw new UserFacingError(`There is a style called “${clean}” already.`)
  }
  return clean
}

/** Saves a format under a name. At most 12 styles: delete one to make room. */
export function addStyle(name: string, format: ShapeFormat, id = `style-${Date.now().toString(36)}-${Math.random().toString(36).slice(2, 7)}`): NamedStyle[] {
  const styles = loadStyles()
  if (styles.length >= MAX_STYLES) throw new UserFacingError(`You have ${MAX_STYLES} saved styles. Delete one to make room.`)
  const next = [...styles, { id, name: cleanName(name, styles), format }]
  write(STYLES_KEY, next)
  return next
}

export function renameStyle(id: string, name: string): NamedStyle[] {
  const styles = loadStyles()
  const clean = cleanName(name, styles, id)
  const next = styles.map((s) => (s.id === id ? { ...s, name: clean } : s))
  write(STYLES_KEY, next)
  return next
}

export function deleteStyle(id: string): NamedStyle[] {
  const next = loadStyles().filter((s) => s.id !== id)
  write(STYLES_KEY, next)
  return next
}

const colourName = (hex: string): string => PALETTE.find((p) => p.hex === hex.toUpperCase())?.name.toLowerCase() ?? `#${hex.toUpperCase()}`

/** A short description for the pane: "navy fill, no outline, Aptos 14 pt bold in white". */
export function describeFormat(format: ShapeFormat): string {
  const parts: string[] = []
  if (format.fill) parts.push(format.fill.kind === 'none' ? 'no fill' : `${colourName(format.fill.colour)} fill`)
  if (format.line) parts.push(format.line.visible ? `${colourName(format.line.colour)} outline ${format.line.weight} pt` : 'no outline')
  const t = format.text
  if (t) {
    const words = [t.name ?? 'mixed fonts', t.size ? `${t.size} pt` : '', t.bold ? 'bold' : '', t.italic ? 'italic' : ''].filter(Boolean)
    parts.push(`${words.join(' ')}${t.colour ? ` in ${colourName(t.colour)}` : ''}`)
  }
  const text = parts.join(', ')
  return text ? `${text.charAt(0).toUpperCase()}${text.slice(1)}` : 'Position and size only'
}
