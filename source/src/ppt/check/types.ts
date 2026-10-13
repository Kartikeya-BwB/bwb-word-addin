// The deck check works on a snapshot of the presentation, so every rule is a pure function that can be tested
// without PowerPoint. Severities are the BwB QA vocabulary the Word check uses: Blocker, Major, Minor, Copyedit.

import type { Severity } from '../../check/types.ts'

export type { Severity }

export interface FontInfo {
  /** Null where the text mixes values. */
  name: string | null
  size: number | null
  /** Upper-case hex without '#'. */
  colour: string | null
  bold: boolean | null
  italic: boolean | null
}

/** A stretch of a shape's text read with getSubstring, because a TextRange has no runs. Offsets index ShapeInfo.text. */
export interface TextPiece {
  start: number
  length: number
  font: FontInfo
}

export interface ShapeInfo {
  id: string
  /** Shape ids from the slide's top level down to this shape: [id] for a shape outside any group. */
  path: string[]
  /** 0 for a shape on the slide, 1 inside a group, 2 inside a group inside a group. */
  level: number
  name: string
  /** PowerPoint's shape type: 'Placeholder', 'TextBox', 'GeometricShape', 'Image', 'Group' and so on. */
  type: string
  /** The placeholder type ('Title', 'CenterTitle', 'Body' and so on), or null for a shape that is not a placeholder. */
  placeholder: string | null
  /** What a placeholder holds ('Image', 'Table' and so on); null when it is empty or PowerPoint cannot say. */
  contains: string | null
  left: number
  top: number
  width: number
  height: number
  /** Tag keys in upper case, as PowerPoint stores them. */
  tags: Record<string, string>
  /** The text, or null for a shape without text (pictures, charts, groups) or one PowerPoint could not read. */
  text: string | null
  /** Formatting of the whole text, null where it was not read. */
  font: FontInfo | null
  /** Where the whole text mixes formatting: uniform stretches, by paragraph and then by word. */
  pieces: TextPiece[]
  fill: { type: string; colour: string | null } | null
  line: { visible: boolean; colour: string | null } | null
  /** Null where it was not read: not a picture, or PowerPoint is older than API 1.10. */
  alt: { title: string; description: string; decorative: boolean } | null
}

export interface SlideInfo {
  id: string
  /** 0-based position in the deck. */
  index: number
  tags: Record<string, string>
  /** The slide layout's name, such as 'Title and Content'. */
  layout: string | null
  /** Every shape, groups before the shapes inside them. */
  shapes: ShapeInfo[]
}

export interface DeckSnapshot {
  slides: SlideInfo[]
  width: number
  height: number
  /** The selected slides, the active one first; empty when PowerPoint cannot say. */
  selected: string[]
  /** What this PowerPoint could read. */
  can: {
    formatting: boolean
    placeholders: boolean
    groups: boolean
    altText: boolean
    pageSize: boolean
    selection: boolean
  }
  /** True when some text mixed so much formatting that part of it was not read piece by piece. */
  partial: boolean
}

export type DeckGroup = 'send' | 'brand' | 'layout' | 'wording' | 'slides'

export interface Span {
  start: number
  length: number
}

export interface Box {
  left: number
  top: number
  width: number
  height: number
}

/** Where a finding is: slide ids and shape ids stay the same while slides move, unlike slide numbers. */
export type DeckLocation =
  | { kind: 'slide'; slideId: string; slide: number }
  | { kind: 'shape'; slideId: string; slide: number; path: string[]; start?: number; end?: number }

export interface ShapeRef {
  slideId: string
  path: string[]
}

export interface TextRef extends ShapeRef {
  /** The shape's whole text when the check ran; the fix stops if it reads differently now. */
  expectedText: string
}

export type DeckFix =
  | (TextRef & { kind: 'replace'; start: number; end: number; match: string; replacement: string })
  | (TextRef & { kind: 'font'; spans: Span[]; font: string })
  | (TextRef & { kind: 'font-size'; spans: Span[]; size: number })
  | (TextRef & { kind: 'text-colour'; spans: Span[]; colour: string })
  | (ShapeRef & { kind: 'fill-colour'; colour: string })
  | (ShapeRef & { kind: 'line-colour'; colour: string })
  | (ShapeRef & { kind: 'align-title'; from: Box; to: Box })
  | (ShapeRef & { kind: 'move-onto-slide'; from: Box; to: Box })
  | (ShapeRef & { kind: 'delete-empty'; textual: boolean })
  | (ShapeRef & { kind: 'delete-shape'; from: Box })
  | { kind: 'delete-tagged'; tag: string; text?: string }
  | { kind: 'update-trackers' }
  | { kind: 'batch'; fixes: DeckFix[] }

export type ShapeFix = Extract<DeckFix, ShapeRef>

export interface DeckFinding {
  id: string
  rule: string
  group: DeckGroup
  severity: Severity
  title: string
  detail?: string
  excerpt?: string
  /** Advisory findings are suggestions: they never stop a deck being ready for QA. */
  advisory?: boolean
  /** The 1-based slide number, for the pane. */
  slide?: number
  location?: DeckLocation
  fix?: DeckFix
  fixLabel?: string
}

export interface DeckStats {
  slides: number
  shapes: number
  words: number
  pictures: number
}

export interface DeckReport {
  findings: DeckFinding[]
  counts: Record<Severity, number>
  /** No blocker or major findings. */
  ready: boolean
  stats: DeckStats
  /** Checks this PowerPoint could not run, in plain words. */
  limits: string[]
}
