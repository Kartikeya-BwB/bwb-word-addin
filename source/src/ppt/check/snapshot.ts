// Reads the presentation for the Pre-QA check, the storyline and the sections in a few batched syncs, however long the
// deck: the slides; their shapes, tags and layouts; each shape's placeholder type, text, font, fill, outline and alt
// text; the shapes inside groups, a level at a time; and, where a shape mixes formatting, its paragraphs and then its
// words, read one piece at a time with getSubstring because a TextRange has no runs.

import { SLIDE } from '../../brand/slideGrid.ts'
import { NEEDS, syncEach } from '../slides/office.ts'
import { requirePpt, runPpt, supportsPpt } from '../run.ts'
import { MEDIA_PLACEHOLDERS, NON_TEXT_CONTENT, needsAltText, placeholderFromName } from './shapes.ts'
import { readBackgrounds, readLinks, readMarkers } from './snapshotExtras.ts'
import { hexColour, splitParagraphs, splitWords } from './text.ts'
import type { DeckSnapshot, FontInfo, ShapeInfo, SlideInfo, Span } from './types.ts'

const SHAPE_PROPS = ['id', 'name', 'type', 'left', 'top', 'width', 'height'].map((p) => `items/${p}`).join(',')
const TAG_PROPS = 'items/key,items/value'
const FONT_PROPS = 'name,size,color,bold,italic'
const TEXT_TYPES = new Set(['GeometricShape', 'TextBox', 'Placeholder', 'Callout', 'Freeform'])
/** Pieces read one by one in a sync, at most; text beyond it keeps only its whole-text formatting. */
const MAX_PIECES = 5000

export interface ReadOptions {
  /** Fonts, colours, fills, outlines and alt text as well as text and tags. The slide tools need only text and tags. */
  formatting?: boolean
  /** Superscript footnote markers (API 1.8). Read with formatting unless set. */
  markers?: boolean
  /** Read only these slides (their index is still their place in the deck). */
  slideIds?: string[]
  /** Read only the current slide. */
  current?: boolean
}

type Can = DeckSnapshot['can']

interface Job {
  shape: PowerPoint.Shape
  info: ShapeInfo
  members: Job[]
  tags: PowerPoint.TagCollection | null
  placeholder: PowerPoint.PlaceholderFormat | null
  group: PowerPoint.ShapeScopedCollection | null
  range: PowerPoint.TextRange | null
  font: PowerPoint.ShapeFont | null
  fill: PowerPoint.ShapeFill | null
  line: PowerPoint.ShapeLineFormat | null
  alt: boolean
  visible: boolean
}

interface SlideJob {
  slide: PowerPoint.Slide
  info: SlideInfo
  shapes: PowerPoint.ShapeCollection | null
  tags: PowerPoint.TagCollection | null
  layout: PowerPoint.SlideLayout | null
  top: Job[]
}

interface PieceJob {
  job: Job
  span: Span
  words: boolean
  font: PowerPoint.ShapeFont | null
}

function tagRecord(items: PowerPoint.Tag[]): Record<string, string> {
  return Object.fromEntries(items.map((t) => [t.key.toUpperCase(), t.value ?? '']))
}

function fontInfo(f: PowerPoint.ShapeFont): FontInfo {
  return {
    name: typeof f.name === 'string' && f.name.trim() ? f.name : null,
    size: typeof f.size === 'number' ? f.size : null,
    colour: hexColour(f.color),
    bold: typeof f.bold === 'boolean' ? f.bold : null,
    italic: typeof f.italic === 'boolean' ? f.italic : null,
  }
}

const mixed = (f: FontInfo): boolean => f.name === null || f.size === null || f.colour === null

function newJob(shape: PowerPoint.Shape, parent: string[], level: number): Job {
  const type = String(shape.type)
  const info: ShapeInfo = {
    id: shape.id, path: [...parent, shape.id], level, name: shape.name ?? '', type,
    placeholder: type === 'Placeholder' ? placeholderFromName(shape.name ?? '') : null, contains: null,
    left: shape.left, top: shape.top, width: shape.width, height: shape.height,
    tags: {}, text: null, font: null, pieces: [], fill: null, line: null, alt: null, visible: null, markers: [],
  }
  return { shape, info, members: [], tags: null, placeholder: null, group: null, range: null, font: null, fill: null, line: null, alt: false, visible: false }
}

function clear(job: Job): void {
  Object.assign(job, { tags: null, placeholder: null, group: null, range: null, font: null, fill: null, line: null, alt: false, visible: false })
}

/** Shapes with text to read: not pictures or charts, and not placeholders made to hold them. */
function textual(info: ShapeInfo): boolean {
  if (!TEXT_TYPES.has(info.type)) return false
  if (info.placeholder !== null && MEDIA_PLACEHOLDERS.has(info.placeholder)) return false
  return !(info.contains !== null && NON_TEXT_CONTENT.has(info.contains))
}

/** Round one for a level of shapes: tags, placeholder types and the shapes inside groups. */
function queueStructure(job: Job, can: Can): void {
  clear(job)
  job.tags = job.shape.tags
  job.tags.load(TAG_PROPS)
  if (job.info.type === 'Placeholder' && can.placeholders) {
    job.placeholder = job.shape.placeholderFormat
    job.placeholder.load('type,containedType')
  }
  if (job.info.type === 'Group' && can.groups) {
    job.group = job.shape.group.shapes
    job.group.load(SHAPE_PROPS)
  }
}

function takeStructure(job: Job): void {
  if (job.tags) job.info.tags = tagRecord(job.tags.items)
  if (job.placeholder) {
    job.info.placeholder = String(job.placeholder.type)
    job.info.contains = job.placeholder.containedType === null ? null : String(job.placeholder.containedType)
  }
  if (job.group) job.members = job.group.items.map((s) => newJob(s, job.info.path, job.info.level + 1))
}

/** Round two: text, and with formatting the font, fill, outline and alt text. */
function queueContent(job: Job, can: Can): void {
  clear(job)
  const { shape, info } = job
  const text = textual(info)
  if (text) {
    job.range = shape.textFrame.textRange
    job.range.load('text')
  }
  if (!can.formatting) return
  if (text) {
    job.font = job.range ? job.range.font : null
    job.font?.load(FONT_PROPS)
    job.fill = shape.fill
    job.fill.load('type,foregroundColor,transparency')
  }
  if (text || info.type === 'Line') {
    job.line = shape.lineFormat
    job.line.load('visible,color')
  }
  if (can.altText && needsAltText(info)) {
    shape.load('altTextTitle,altTextDescription,isDecorative')
    job.alt = true
  }
  if (can.hidden) {
    shape.load('visible')
    job.visible = true
  }
}

function takeContent(job: Job): void {
  const { info, shape } = job
  if (job.range) info.text = job.range.text ?? ''
  if (job.font) info.font = fontInfo(job.font)
  if (job.fill) {
    const transparency = typeof job.fill.transparency === 'number' ? job.fill.transparency : 0
    info.fill = { type: String(job.fill.type), colour: hexColour(job.fill.foregroundColor), ...(transparency > 0 ? { transparency } : {}) }
  }
  if (job.line) info.line = { visible: job.line.visible === true, colour: hexColour(job.line.color) }
  if (job.alt) info.alt = { title: shape.altTextTitle ?? '', description: shape.altTextDescription ?? '', decorative: shape.isDecorative === true }
  if (job.visible) info.visible = shape.visible !== false
}

/** Splits mixed text into paragraphs, and paragraphs that still mix formatting into words. Returns false if capped. */
async function readPieces(context: PowerPoint.RequestContext, jobs: Job[]): Promise<boolean> {
  let complete = true
  let pending: PieceJob[] = []
  for (const job of jobs) {
    const { info } = job
    if (!job.range || !info.font || !info.text || !mixed(info.font)) continue
    const paragraphs = splitParagraphs(info.text).filter((p) => p.text.trim() !== '')
    const only = paragraphs.length === 1 ? paragraphs[0] : undefined
    if (only && only.length === info.text.length) {
      pending.push(...splitWords(info.text).map((span) => ({ job, span, words: true, font: null })))
    } else {
      pending.push(...paragraphs.map((p) => ({ job, span: { start: p.start, length: p.length }, words: false, font: null })))
    }
  }
  while (pending.length) {
    if (pending.length > MAX_PIECES) {
      complete = false
      pending = pending.slice(0, MAX_PIECES)
    }
    const failed = await syncEach(context, pending, (p) => {
      p.font = null
      const range = p.job.range
      if (!range) return
      p.font = range.getSubstring(p.span.start, p.span.length).font
      p.font.load(FONT_PROPS)
    })
    const next: PieceJob[] = []
    for (const p of pending) {
      if (failed.has(p) || !p.font) continue
      const font = fontInfo(p.font)
      const text = p.job.info.text ?? ''
      if (mixed(font) && !p.words) {
        next.push(...splitWords(text, p.span.start, p.span.length).map((span) => ({ job: p.job, span, words: true, font: null })))
      } else {
        p.job.info.pieces.push({ start: p.span.start, length: p.span.length, font })
      }
    }
    pending = next
  }
  for (const job of jobs) job.info.pieces.sort((a, b) => a.start - b.start)
  return complete
}

function flatten(jobs: Job[]): ShapeInfo[] {
  return jobs.flatMap((j) => [j.info, ...flatten(j.members)])
}

/** Reads the deck inside an existing PowerPoint.run. */
export async function readDeck(context: PowerPoint.RequestContext, options: ReadOptions = {}): Promise<DeckSnapshot> {
  const formatting = options.formatting ?? true
  const can: Can = {
    formatting,
    placeholders: supportsPpt(NEEDS.move),
    groups: supportsPpt(NEEDS.move),
    altText: formatting && supportsPpt(NEEDS.export),
    pageSize: supportsPpt(NEEDS.export),
    selection: supportsPpt(NEEDS.current),
    hidden: formatting && supportsPpt(NEEDS.export),
    background: formatting && supportsPpt(NEEDS.export),
    links: formatting && supportsPpt(NEEDS.links),
    markers: (options.markers ?? formatting) && supportsPpt(NEEDS.move),
  }
  const presentation = context.presentation
  const slides = presentation.slides
  slides.load('items/id')
  const selected = can.selection ? presentation.getSelectedSlides() : null
  selected?.load('items/id')
  const page = can.pageSize ? presentation.pageSetup : null
  page?.load('slideWidth,slideHeight')
  await context.sync()

  const active = selected?.items[0]?.id
  const wanted = options.current ? new Set(active ? [active] : []) : options.slideIds ? new Set(options.slideIds) : null
  const slideJobs: SlideJob[] = slides.items
    .map((slide, index): SlideJob => ({
      slide, shapes: null, tags: null, layout: null, top: [],
      info: { id: slide.id, index, tags: {}, layout: null, shapes: [], background: null, links: [] },
    }))
    .filter((s) => !wanted || wanted.has(s.info.id))
  const failedSlides = await syncEach(context, slideJobs, (s) => {
    s.shapes = s.slide.shapes
    s.shapes.load(SHAPE_PROPS)
    s.tags = s.slide.tags
    s.tags.load(TAG_PROPS)
    s.layout = s.slide.layout
    s.layout.load('name')
  })
  for (const s of slideJobs) {
    if (failedSlides.has(s) || !s.shapes || !s.tags || !s.layout) continue
    s.info.tags = tagRecord(s.tags.items)
    s.info.layout = s.layout.name || null
    s.top = s.shapes.items.map((shape) => newJob(shape, [], 0))
  }

  const all: Job[] = []
  let level = slideJobs.flatMap((s) => s.top)
  while (level.length) {
    all.push(...level)
    const failedStructure = await syncEach(context, level, (job) => queueStructure(job, can))
    level.forEach((job) => (failedStructure.has(job) ? clear(job) : takeStructure(job)))
    const failedContent = await syncEach(context, level, (job) => queueContent(job, can))
    level.forEach((job) => (failedContent.has(job) ? clear(job) : takeContent(job)))
    level = level.flatMap((job) => job.members)
  }
  const piecesRead = can.formatting ? await readPieces(context, all) : true
  const read = slideJobs.filter((s) => !failedSlides.has(s))
  if (can.background) await readBackgrounds(context, read)
  if (can.links) await readLinks(context, read, supportsPpt(NEEDS.export))
  const markersRead = can.markers ? await readMarkers(context, all) : true
  const complete = piecesRead && markersRead

  for (const s of slideJobs) s.info.shapes = flatten(s.top)
  return {
    slides: slideJobs.map((s) => s.info),
    width: page?.slideWidth ?? SLIDE.width,
    height: page?.slideHeight ?? SLIDE.height,
    selected: selected ? selected.items.map((s) => s.id) : [],
    can,
    partial: !complete,
  }
}

/** Reads the open presentation. Needs PowerPoint API 1.4; newer PowerPoint adds placeholder types, groups and alt text. */
export async function takeDeckSnapshot(options: ReadOptions = {}): Promise<DeckSnapshot> {
  requirePpt(NEEDS.read, 'Reading the slides')
  return runPpt((context) => readDeck(context, options))
}
