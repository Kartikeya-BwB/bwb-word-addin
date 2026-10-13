// The Pre-QA check works on a snapshot of the document, so every rule is a pure function that can be
// tested without Word. Severities follow the BwB QA vocabulary: Blocker, Major, Minor, Copyedit.

import type { BuiltIn } from '../brand/styles.ts'
import type { AcronymEntry } from '../rules/acronyms.ts'
import type { HyphenVariant } from '../rules/hyphenation.ts'
import type { DocInfo } from '../word/properties.ts'

export interface FontInfo {
  name: string | null
  size: number | null
  color: string | null
  highlight: string | null
  bold: boolean | null
}

export interface ParagraphInfo {
  index: number
  text: string
  style: string
  styleBuiltIn: string
  isListItem: boolean
  inTable: boolean
  font: FontInfo
}

export interface NeighbourParagraph {
  text: string
  style: string
}

export interface TableInfo {
  index: number
  style: string
  rows: number
  firstCell: string
  after: NeighbourParagraph | null
  before: NeighbourParagraph | null
}

export interface PictureInfo {
  index: number
  altTitle: string
  altDescription: string
}

export interface ControlInfo {
  index: number
  tag: string
  title: string
  text: string
}

export interface HeaderFooterInfo {
  section: number
  kind: 'header' | 'footer'
  text: string
  pictures: number
}

/** Run-level formatting found by reading the body's WordprocessingML, per paragraph index. */
export interface RunFormatting {
  paragraph: number | null
  fonts: string[]
  colours: string[]
  highlights: string[]
  smallestHalfPoints: number | null
  /** Inline pictures (icons, figures), which Word leaves out of the paragraph's text. */
  pictures: number
}

export interface DocSnapshot {
  paragraphs: ParagraphInfo[]
  tables: TableInfo[]
  pictures: PictureInfo[]
  controls: ControlInfo[]
  headersFooters: HeaderFooterInfo[]
  /** Null where this Word does not offer the API (fields need WordApi 1.5, comments 1.4, tracked changes 1.6). */
  fieldCodes: string[] | null
  comments: number | null
  trackedChanges: number | null
  formatting: RunFormatting[] | null
  info: DocInfo
  fileName: string | null
  bwbStyles: boolean | null
}

export type Severity = 'blocker' | 'major' | 'minor' | 'copyedit'
export const SEVERITIES: Severity[] = ['blocker', 'major', 'minor', 'copyedit']

export type CheckGroup = 'send' | 'writing' | 'structure' | 'tables' | 'brand' | 'consistency'

export type Location =
  | { kind: 'paragraph'; paragraph: number; expectedText: string; start?: number; end?: number }
  | { kind: 'table'; table: number }
  | { kind: 'picture'; picture: number }
  | { kind: 'control'; control: number; tag: string }

export type FixAction =
  | { kind: 'replace'; paragraph: number; expectedText: string; start: number; end: number; match: string; replacement: string }
  | { kind: 'paragraph-style'; paragraph: number; expectedText: string; style: string; builtIn?: BuiltIn }
  | { kind: 'paragraph-font'; paragraph: number; expectedText: string; font: string }
  | { kind: 'remove-highlight'; paragraph: number; expectedText: string }
  | { kind: 'delete-paragraph'; paragraph: number; expectedText: string }
  | { kind: 'table-style'; table: number; style: string }
  | { kind: 'properties'; values: Partial<DocInfo> }
  | { kind: 'remove-controls'; tag: string }
  | { kind: 'remove-sticker' }
  | { kind: 'insert-disclaimer' }
  | { kind: 'add-styles' }

export interface Finding {
  id: string
  rule: string
  group: CheckGroup
  severity: Severity
  title: string
  detail?: string
  excerpt?: string
  /** Advisory findings are suggestions: they never stop a document being ready for QA. */
  advisory?: boolean
  location?: Location
  fix?: FixAction
  fixLabel?: string
}

export interface CheckStats {
  paragraphs: number
  words: number
  headings: number
  tables: number
  pictures: number
}

export interface CheckReport {
  findings: Finding[]
  counts: Record<Severity, number>
  /** No blocker or major findings. */
  ready: boolean
  acronyms: AcronymEntry[]
  hyphenation: HyphenVariant[]
  stats: CheckStats
}
