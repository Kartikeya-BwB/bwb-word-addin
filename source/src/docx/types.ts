/**
 * The BwB content model (references/content-schema.md in the BwB skill), as the TypeScript builder takes it.
 *
 * Text fields accept the inline mark-up: **bold**, *italic*, [label](https://...), [label](#bookmark), ^[footnote],
 * {icon:name} or {icon:name:colour}, escapes \* \[ \] \^ \\, and the add-in token {field:Label}.
 * Field names follow the JSON schema (snake_case) so content written for the Python builder works unchanged,
 * except that pictures are passed as bytes rather than file paths.
 */

/** The five approved icon treatments: green, navy and charcoal on light backgrounds; white and lime on navy panels. */
export type IconColour = 'green' | 'navy' | 'charcoal' | 'white' | 'lime'

/** Draws Lucide icons in the BwB colours. The add-in supplies one for the browser; tests use lucide-static and resvg. */
export interface IconRenderer {
  /** The current Lucide name for a name or an older alias (iconPolicy.resolveIconName does the matching), or null. */
  resolve(name: string): string | null
  /** SVG markup of a resolved icon, recoloured (iconPolicy.recolourSvg), or null when it cannot be drawn. */
  svg(name: string, colour: IconColour): string | null
  /** A square PNG of a resolved icon, px wide, or null when it cannot be drawn. */
  png(name: string, colour: IconColour, px: number): Promise<Uint8Array | null>
  /** Optional: every icon name and older alias, for "did you mean" suggestions in warnings. */
  names?(): readonly string[]
}

/** A picture passed as bytes. Word embeds PNG, JPEG and GIF inline. */
export interface PictureData {
  data: Uint8Array
  ext: 'png' | 'jpg' | 'gif'
  widthPx: number
  heightPx: number
}

export interface ListItemObject {
  text: string
  /** 0 (default) to 2; deeper levels are set at level 2. */
  level?: number
}
export type ListItem = string | ListItemObject

export type BulletStyle = 'green-square' | 'navy-circle' | 'green-circle' | 'charcoal-dash' | 'arrow' | 'lime-navy'

export interface HeadingBlock {
  type: 'heading'
  text: string
  level?: number
  numbered?: boolean
  in_toc?: boolean
  /** Subtitle of the section divider inserted before this heading when meta.dividers is 'h1'. */
  subtitle?: string
  divider_variant?: 'dark' | 'light'
}
export interface ParagraphBlock {
  type: 'paragraph'
  text: string
}
export interface LeadBlock {
  type: 'lead'
  text: string
}
export interface AccentBlock {
  type: 'accent'
  text: string
}
export interface SourceBlock {
  type: 'source'
  text: string
}
export interface BulletsBlock {
  type: 'bullets'
  items: ListItem[]
  style?: BulletStyle
}
export interface NumberedBlock {
  type: 'numbered'
  items: ListItem[]
}
export interface PullQuoteBlock {
  type: 'pull_quote'
  text: string
  attribution?: string
}
export interface PageBreakBlock {
  type: 'page_break'
}

/** A table cell: text with mark-up, a number, empty (null), or a list of green-circle bullets. */
export type TableCell = string | number | null | { bullets: string[] }
export type ColumnAlign = 'center' | 'right' | 'left' | null

export interface TableBlock {
  type: 'table'
  header?: string[] | null
  rows: TableCell[][]
  style?: 'standard' | 'comparison' | 'status'
  /** Relative column weights; computed from the contents when absent. */
  widths?: number[]
  align?: ColumnAlign[]
  total_row?: boolean
  /** Caption text only; "Table N." is added with a SEQ field. */
  caption?: string
  source?: string
  alt?: string
}

export interface FigureBlock extends PictureData {
  type: 'figure'
  caption?: string
  source?: string
  alt?: string
  width_cm?: number
}

export interface StatItem {
  value: string
  label?: string
  icon?: string
  icon_colour?: string
}
export interface StatsBlock {
  type: 'stats'
  items: StatItem[]
  /** Index of the item shown in lime. */
  hero?: number | null
}

export interface SnapshotItem {
  eyebrow?: string
  value?: string
  note?: string
  icon?: string
  icon_colour?: string
}
export interface SnapshotBlock {
  type: 'snapshot'
  items: SnapshotItem[]
  columns?: number
}

export type KeyValueRows = Array<[string, string]> | Record<string, string>
export interface KeyValuesBlock {
  type: 'key_values'
  rows: KeyValueRows
}

export interface CalloutBlock {
  type: 'callout'
  title?: string
  text?: string | string[]
  bullets?: string[]
  variant?: 'green' | 'charcoal'
  icon?: string
}

export interface CaseStudyStat {
  value: string
  label?: string
}
export interface CaseStudyBlock {
  type: 'case_study'
  title: string
  meta?: string
  body: string | string[]
  stats?: CaseStudyStat[]
  source?: string
  icon?: string
}

export interface IconGridItem {
  icon?: string
  title?: string
  text?: string | string[]
  icon_colour?: string
}
export interface IconGridBlock {
  type: 'icon_grid'
  items: IconGridItem[]
  columns?: number
  variant?: 'light' | 'navy'
  /** Index of the tile whose icon is lime (navy variant). */
  hero?: number | null
}

export interface DividerBlock {
  type: 'divider'
  number?: string | number
  title: string
  subtitle?: string
  variant?: 'dark' | 'light'
}

// ---- add-in extensions (not in the Python builder) ----

/** A writing prompt for a new document: a rich text content control (tag BwBGuidance) around one BwBLead paragraph. */
export interface GuidanceBlock {
  type: 'guidance'
  text: string
}

/** The standard BwB disclaimer page, exactly as the builder ends a document (page break before the heading). */
export interface DisclaimerBlock {
  type: 'disclaimer'
}

/** A table of contents field with no cached entries (Word builds them when fields are updated), then a page break. */
export interface TocBlock {
  type: 'toc'
  /** Heading levels listed, 1 to 9; default 2. */
  levels?: number
  /** Default "Table of Contents". */
  title?: string
}

/** A caption paragraph on its own, numbered with a SEQ field like the captions of tables and figures. */
export interface CaptionBlock {
  type: 'caption'
  kind: 'Table' | 'Figure'
  text: string
  source?: string
}

export type Block =
  | HeadingBlock
  | ParagraphBlock
  | LeadBlock
  | AccentBlock
  | SourceBlock
  | BulletsBlock
  | NumberedBlock
  | PullQuoteBlock
  | PageBreakBlock
  | TableBlock
  | FigureBlock
  | StatsBlock
  | SnapshotBlock
  | KeyValuesBlock
  | CalloutBlock
  | CaseStudyBlock
  | IconGridBlock
  | DividerBlock
  | GuidanceBlock
  | DisclaimerBlock
  | TocBlock
  | CaptionBlock

export interface ContentMeta {
  /** Default "Untitled document". */
  title?: string
  subtitle?: string
  tagline?: string
  /** 'report' (default): cover and, for five or more headings, contents. 'memo': title block and details. */
  format?: 'report' | 'memo'
  /** "5 October 2026"; default today. */
  date?: string
  /** Default "Bankers without Boundaries". */
  prepared_by?: string
  meta_left?: string
  /** Default "bwb.earth"; an empty string leaves the right side blank. */
  meta_right?: string
  /** A photograph for the cover instead of the BwB cover photo. */
  cover_image?: PictureData
  details?: KeyValueRows
  toc?: boolean | 'auto'
  toc_levels?: number
  toc_title?: string
  numbering?: 'auto' | 'all' | 'none'
  dividers?: 'none' | 'h1'
  h1_page_break?: boolean | 'auto'
  /** Default true. */
  disclaimer?: boolean
  disclaimer_paragraphs?: string[]
  disclaimer_note?: string
  subject?: string
  keywords?: string
  description?: string
  /** Add-in extension: Office custom document properties (docProps/custom.xml), such as "BwB document type". */
  custom_properties?: Record<string, string>
}

export interface Content {
  meta?: ContentMeta
  blocks: Block[]
}

export interface BuildAssets {
  /** public/templates/bwb-base.docx */
  base: Uint8Array
  /** public/templates/bwb-cover.jpg; required for reports unless meta.cover_image is given. */
  cover?: Uint8Array
  icons?: IconRenderer
  /** Date for the default meta.date and the document properties; default now. */
  now?: Date
}

export interface BlockOptions {
  /** Number shown by the first new caption of each kind until fields are updated; default 1 (the next number). */
  seqStart?: { Table?: number; Figure?: number }
  icons?: IconRenderer
}

export interface BuildResult {
  docx: Uint8Array
  warnings: string[]
}

export interface BlockResult {
  ooxml: string
  warnings: string[]
}
