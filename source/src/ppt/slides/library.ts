// The BwB slide library: ready-made slides in public/templates/bwb-slides.pptx, listed with a thumbnail each in
// public/assets/slides/index.json. A slide goes in after the current one with insertSlidesFromBase64; the .pptx is
// fetched once and kept as base64 for the session.

import { TAGS } from '../../brand/slideGrid.ts'
import { UserFacingError } from '../../word/run.ts'
import { NEEDS } from './office.ts'
import { requirePpt, runPpt, supportsPpt } from '../run.ts'

export type SlideFormatting = 'KeepSourceFormatting' | 'UseDestinationTheme'

export const FORMATTING_OPTIONS: Array<{ value: SlideFormatting; label: string; hint: string }> = [
  { value: 'KeepSourceFormatting', label: 'Keep the BwB look', hint: 'The slide keeps the BwB template’s fonts, colours and layout.' },
  { value: 'UseDestinationTheme', label: 'Match this deck', hint: 'The slide takes this deck’s theme, for a client’s template.' },
]

export interface LibrarySlide {
  id: string
  title: string
  category: string
  description: string
  tags: string[]
  /** The slide's id in the library .pptx, such as "256#1234567". */
  slideId: string
  /** Thumbnail path relative to the add-in's root. */
  thumb: string
}

export interface SlideLibraryIndex {
  version: 1
  /** The library .pptx, relative to the add-in's root. */
  file: string
  slides: LibrarySlide[]
}

export class SlideLibraryError extends UserFacingError {}

const isString = (value: unknown): value is string => typeof value === 'string'

function field(entry: object, key: string): unknown {
  return Object.prototype.hasOwnProperty.call(entry, key) ? (entry as Record<string, unknown>)[key] : undefined
}

function parseSlide(entry: unknown): LibrarySlide | null {
  if (typeof entry !== 'object' || entry === null) return null
  const [id, title, category, slideId, thumb] = ['id', 'title', 'category', 'slideId', 'thumb'].map((k) => field(entry, k))
  const description = field(entry, 'description')
  const tags = field(entry, 'tags')
  if (!isString(id) || !id.trim() || !isString(title) || !isString(category) || !isString(slideId) || !slideId.trim() || !isString(thumb)) return null
  return {
    id, title, category, slideId, thumb,
    description: isString(description) ? description : '',
    tags: Array.isArray(tags) ? tags.filter(isString).map((t) => t.toLowerCase()) : [],
  }
}

/** Checks the index the library build writes. Entries without an id, title, category, slide id or thumbnail are left out. */
export function parseSlideIndex(data: unknown): SlideLibraryIndex {
  if (typeof data !== 'object' || data === null) throw new SlideLibraryError('The slide library index is not in the expected format.')
  const version = field(data, 'version')
  const file = field(data, 'file')
  const slides = field(data, 'slides')
  if (version !== 1 || !isString(file) || !file.trim() || !Array.isArray(slides)) {
    throw new SlideLibraryError('The slide library index is not in the expected format.')
  }
  const seen = new Set<string>()
  const parsed: LibrarySlide[] = []
  for (const entry of slides) {
    const slide = parseSlide(entry)
    if (!slide || seen.has(slide.id)) continue
    seen.add(slide.id)
    parsed.push(slide)
  }
  if (!parsed.length) throw new SlideLibraryError('The slide library index lists no slides.')
  return { version: 1, file, slides: parsed }
}

/** 'key-figures' reads 'Key figures'. */
export function categoryLabel(category: string): string {
  const words = category.replace(/[-_]+/g, ' ').trim()
  return words ? words.charAt(0).toUpperCase() + words.slice(1).toLowerCase() : 'Other'
}

/** The categories in the order they first appear, with how many slides each has. */
export function slideCategories(index: SlideLibraryIndex): Array<{ id: string; label: string; count: number }> {
  const counts = new Map<string, number>()
  for (const s of index.slides) counts.set(s.category, (counts.get(s.category) ?? 0) + 1)
  return [...counts].map(([id, count]) => ({ id, label: categoryLabel(id), count }))
}

/** Ranks slides for a query by tags, title, category and description; with no query, keeps the catalogue order. */
export function searchSlides(index: SlideLibraryIndex, query: string, category = ''): LibrarySlide[] {
  const pool = category ? index.slides.filter((s) => s.category === category) : index.slides
  const words = query.toLowerCase().split(/[\s,]+/).filter(Boolean)
  if (!words.length) return pool
  const scored = pool.map((slide, order) => {
    const title = slide.title.toLowerCase()
    const cat = slide.category.toLowerCase()
    const description = slide.description.toLowerCase()
    let score = 0
    for (const w of words) {
      if (slide.tags.includes(w)) score += 3
      else if (slide.tags.some((t) => t.startsWith(w) || t.includes(` ${w}`))) score += 2
      if (title.includes(w)) score += 2
      if (cat.includes(w)) score += 1
      if (description.includes(w)) score += 1
    }
    return { slide, score, order }
  })
  return scored.filter((s) => s.score > 0).sort((a, b) => b.score - a.score || a.order - b.order).map((s) => s.slide)
}

/** A library slide by id ('agenda'), by id prefix ('agenda-numbered'), or else the first of that category. */
export function findLibrarySlide(index: SlideLibraryIndex, key: string): LibrarySlide | undefined {
  const k = key.toLowerCase()
  return (
    index.slides.find((s) => s.id.toLowerCase() === k) ??
    index.slides.find((s) => s.id.toLowerCase().startsWith(`${k}-`)) ??
    index.slides.find((s) => s.category.toLowerCase() === k)
  )
}

export function bytesToBase64(bytes: Uint8Array): string {
  let binary = ''
  const chunk = 0x8000
  for (let i = 0; i < bytes.length; i += chunk) binary += String.fromCharCode(...bytes.subarray(i, i + chunk))
  return btoa(binary)
}

let indexCache: Promise<SlideLibraryIndex> | undefined
const fileCache = new Map<string, Promise<string>>()

/** Forgets the cached index and .pptx, for tests. */
export function clearSlideLibraryCache(): void {
  indexCache = undefined
  fileCache.clear()
}

export function loadSlideIndex(url = './assets/slides/index.json'): Promise<SlideLibraryIndex> {
  if (!indexCache) {
    const loading = fetch(url).then(async (r) => {
      if (!r.ok) throw new SlideLibraryError(`The slide library could not be loaded (${r.status}).`)
      return parseSlideIndex(await r.json())
    })
    // A failure is not kept, so the next attempt fetches again.
    loading.catch(() => {
      if (indexCache === loading) indexCache = undefined
    })
    indexCache = loading
  }
  return indexCache
}

/** The library .pptx as base64, fetched once per session. */
export function libraryBase64(file: string): Promise<string> {
  const cached = fileCache.get(file)
  if (cached) return cached
  const loading = fetch(`./${file}`).then(async (r) => {
    if (!r.ok) throw new SlideLibraryError(`The slide library file could not be loaded (${r.status}).`)
    return bytesToBase64(new Uint8Array(await r.arrayBuffer()))
  })
  loading.catch(() => {
    if (fileCache.get(file) === loading) fileCache.delete(file)
  })
  fileCache.set(file, loading)
  return loading
}

export interface InsertedSlides {
  /** The new slides' ids in this deck, in order. */
  slideIds: string[]
  /** True when they went in after the current slide; false when PowerPoint could not say which slide is current. */
  afterCurrent: boolean
}

/** Inserts library slides after the current slide (or at the end), selects the first and returns the new ids. */
export async function insertLibrarySlides(index: SlideLibraryIndex, slides: LibrarySlide[], formatting: SlideFormatting = 'KeepSourceFormatting'): Promise<InsertedSlides> {
  requirePpt(NEEDS.tags, 'Inserting slides from the BwB library')
  const base64 = await libraryBase64(index.file)
  const current = supportsPpt(NEEDS.current)
  return runPpt(async (context) => {
    const presentation = context.presentation
    const before = presentation.slides
    before.load('items/id')
    const selected = current ? presentation.getSelectedSlides() : null
    selected?.load('items/id')
    await context.sync()
    const known = new Set(before.items.map((s) => s.id))
    const target = selected?.items[0]?.id ?? before.items[before.items.length - 1]?.id
    presentation.insertSlidesFromBase64(base64, {
      formatting,
      sourceSlideIds: slides.map((s) => s.slideId),
      ...(target ? { targetSlideId: target } : {}),
    })
    await context.sync()
    const after = presentation.slides
    after.load('items/id')
    await context.sync()
    const slideIds = after.items.map((s) => s.id).filter((id) => !known.has(id))
    const first = slideIds[0]
    if (!first) throw new SlideLibraryError('PowerPoint did not insert the slide. Try again, or copy it from the BwB slide library file.')
    // The slides arrive in the order asked for. Tagging each with its library id tells the slide tools and the deck
    // check what it is (a divider, the agenda) whatever its layout is called in this deck.
    slideIds.forEach((id, i) => {
      const entry = slides[i]
      if (entry) presentation.slides.getItem(id).tags.add(TAGS.layout, entry.id)
    })
    if (current) presentation.setSelectedSlides([first])
    await context.sync()
    return { slideIds, afterCurrent: Boolean(selected?.items[0]) }
  })
}

/** Inserts one library slide after the current slide. */
export async function insertLibrarySlide(slide: LibrarySlide, formatting: SlideFormatting = 'KeepSourceFormatting', index?: SlideLibraryIndex): Promise<InsertedSlides> {
  return insertLibrarySlides(index ?? (await loadSlideIndex()), [slide], formatting)
}
