// What a shape or slide is, from its type, placeholder, tags, name and position: titles, sources, BwB furniture,
// pictures, and title, agenda and divider slides. The deck check, storyline and sections all decide the same way.

import { SOURCE_AREA, TAGS } from '../../brand/slideGrid.ts'
import { oneLine } from './text.ts'
import type { ShapeInfo, SlideInfo } from './types.ts'

export const TITLE_PLACEHOLDERS = new Set(['Title', 'CenterTitle', 'VerticalTitle'])
export const TEXT_PLACEHOLDERS = new Set(['Title', 'CenterTitle', 'VerticalTitle', 'Subtitle', 'Body', 'VerticalBody', 'Content', 'VerticalContent'])
export const BODY_PLACEHOLDERS = new Set(['Body', 'VerticalBody', 'Content', 'VerticalContent'])
/** Placeholders for slide furniture, which the layout fills in. */
export const FOOTER_PLACEHOLDERS = new Set(['Footer', 'SlideNumber', 'Date', 'Header'])
/** Placeholders for pictures, charts, tables and media rather than text. */
export const MEDIA_PLACEHOLDERS = new Set(['Picture', 'Chart', 'Table', 'SmartArt', 'Media', 'OnlinePicture', 'Cameo'])
/** Shapes that do not hold text of their own. */
export const NON_TEXT_CONTENT = new Set(['Image', 'Table', 'Chart', 'SmartArt', 'Media', 'Graphic', 'Ole', 'ContentApp', 'Diagram', 'Model3D', 'Ink', 'Group', 'Line'])

/** The value of a tag, matching the key in any case. */
export function tagValue(owner: { tags: Record<string, string> }, key: string): string | undefined {
  return owner.tags[key.toUpperCase()]
}

export function hasTag(owner: { tags: Record<string, string> }, key: string): boolean {
  return tagValue(owner, key) !== undefined
}

/** How a tagged shape reads in a list: its tag's value (a sticker's label), else its text, in capitals. */
export function tagLabel(s: ShapeInfo, key: string): string {
  return (tagValue(s, key)?.trim() || oneLine(s.text ?? '')).toUpperCase()
}

/** A placeholder's type from the name PowerPoint gives it ("Title 1", "Content Placeholder 2"), for PowerPoint before API 1.8. */
export function placeholderFromName(name: string): string {
  const n = name.trim().toLowerCase()
  if (/^(?:center |centre )?title\b/.test(n)) return 'Title'
  if (/^subtitle\b/.test(n)) return 'Subtitle'
  if (/^slide number\b/.test(n)) return 'SlideNumber'
  if (/^footer\b/.test(n)) return 'Footer'
  if (/^date\b/.test(n)) return 'Date'
  if (/^header\b/.test(n)) return 'Header'
  if (/^picture\b/.test(n)) return 'Picture'
  if (/^chart\b/.test(n)) return 'Chart'
  if (/^table\b/.test(n)) return 'Table'
  if (/^(?:content|text|body)\b/.test(n)) return 'Body'
  return 'Unsupported'
}

export function isTitleShape(s: ShapeInfo): boolean {
  return s.placeholder !== null && TITLE_PLACEHOLDERS.has(s.placeholder)
}

export function isFooterShape(s: ShapeInfo): boolean {
  return s.placeholder !== null && FOOTER_PLACEHOLDERS.has(s.placeholder)
}

/** Shapes the add-in puts on slides and manages: trackers, stickers and review notes. */
export function isBwbFurniture(s: ShapeInfo): boolean {
  return hasTag(s, TAGS.tracker) || hasTag(s, TAGS.sticker) || hasTag(s, TAGS.note)
}

/** Source lines and footnotes, which may be set at 8 pt: named or tagged so, or sitting in the source band. */
export function isSourceShape(s: ShapeInfo): boolean {
  if (/source|footnote/i.test(s.name)) return true
  if (/source|footnote/i.test(tagValue(s, TAGS.element) ?? '')) return true
  return s.top >= SOURCE_AREA.top - 8 && s.height <= SOURCE_AREA.height * 2
}

export function isPicture(s: ShapeInfo): boolean {
  return s.type === 'Image' || (s.type === 'Placeholder' && s.contains === 'Image')
}

/** An SVG graphic, such as an icon, which PowerPoint reports as type Graphic. */
export function isGraphic(s: ShapeInfo): boolean {
  return s.type === 'Graphic' || (s.type === 'Placeholder' && s.contains === 'Graphic')
}

/** Shapes that need alt text: pictures and graphics. PowerPoint keeps alt text, or a decorative mark, on both. */
export function needsAltText(s: ShapeInfo): boolean {
  return isPicture(s) || isGraphic(s)
}

export function hasText(s: ShapeInfo): boolean {
  return s.text !== null && s.text.trim() !== ''
}

/** Quotations keep their words as spoken: tagged or named as quotes, or set in Candara italic like the BwB quote. */
export function isQuoteShape(s: ShapeInfo): boolean {
  if (/quote/i.test(tagValue(s, TAGS.element) ?? '') || /quote/i.test(s.name)) return true
  return !isTitleShape(s) && s.font !== null && s.font.italic === true && (s.font.name ?? '').toLowerCase() === 'candara'
}

/** About two lines of a 24 pt Candara action title across the BwB title area. */
export const LONG_TITLE = 90

export type SlideKind = 'title' | 'agenda' | 'divider' | 'summary' | 'closing' | 'content'

const LAYOUT_TAG_KINDS: Array<[RegExp, SlideKind]> = [
  [/^(?:title|cover)(?:-(?:slide|photo|image|navy))?$/, 'title'],
  [/^(?:agenda|contents)\b/, 'agenda'],
  [/^(?:section-)?(?:divider|section)\b/, 'divider'],
  [/^(?:executive-)?summary\b/, 'summary'],
  [/^(?:closing|end|contact|back|disclaimer)\b/, 'closing'],
]

/** Layout names in the BwB template ("BwB Section divider") and PowerPoint's defaults ("Section Header"). */
const LAYOUT_NAME_KINDS: Array<[RegExp, SlideKind]> = [
  [/section header|divider/, 'divider'],
  [/^(?:bwb )?title(?: slide)?$|cover/, 'title'],
  [/closing|thank|end slide/, 'closing'],
]

const TITLE_KINDS: Array<[RegExp, SlideKind]> = [
  [/^(?:agenda|contents)$/, 'agenda'],
  [/^(?:executive summary|summary|key messages)$/, 'summary'],
  [/^thank you\b/, 'closing'],
]

/**
 * What a slide is for: from its BWB_LAYOUT tag (the slide tools tag library slides with their library id), else its
 * layout's name, a centred title, or the title's words.
 */
export function slideKind(slide: SlideInfo): SlideKind {
  const tag = (tagValue(slide, TAGS.layout) ?? '').trim().toLowerCase()
  for (const [re, kind] of LAYOUT_TAG_KINDS) if (re.test(tag)) return kind
  if (tag) return 'content'
  const layout = (slide.layout ?? '').trim().toLowerCase()
  for (const [re, kind] of LAYOUT_NAME_KINDS) if (re.test(layout)) return kind
  if (slide.shapes.some((s) => s.placeholder === 'CenterTitle')) return 'title'
  const title = slideTitle(slide)?.text.toLowerCase() ?? ''
  for (const [re, kind] of TITLE_KINDS) if (re.test(title)) return kind
  return 'content'
}

export interface SlideTitle {
  text: string
  shape: ShapeInfo
  /** True when the title is in a title placeholder rather than the top-most text on the slide. */
  fromPlaceholder: boolean
}

const topFirst = (a: ShapeInfo, b: ShapeInfo): number => a.top - b.top || a.left - b.left

/** The slide's title: its title placeholder's text, or else the top-most text, leaving out furniture and sources. */
export function slideTitle(slide: SlideInfo): SlideTitle | null {
  const placeholder = slide.shapes.filter((s) => isTitleShape(s) && hasText(s)).sort(topFirst)[0]
  if (placeholder) return { text: oneLine(placeholder.text ?? ''), shape: placeholder, fromPlaceholder: true }
  const text = slide.shapes
    .filter((s) => hasText(s) && !isBwbFurniture(s) && !isFooterShape(s) && !isSourceShape(s))
    .sort(topFirst)[0]
  return text ? { text: oneLine(text.text ?? ''), shape: text, fromPlaceholder: false } : null
}

/** Every shape with text that the author wrote: not furniture, not footers. */
export function authoredText(slide: SlideInfo): ShapeInfo[] {
  return slide.shapes.filter((s) => hasText(s) && !isBwbFurniture(s) && !isFooterShape(s))
}
