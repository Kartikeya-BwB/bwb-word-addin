// The icon browser's libraries: Lucide (the BwB default), the BwB sector set, Tabler, flags and maps. Each loads only
// when it is opened or searched. The data comes from scripts/make-icon-index.mjs and scripts/make-icon-libraries.mjs.
// Search ranks the way searchIcons in icons.ts does, so Lucide results come out in the same order.

import type { IconIndex, IconNode } from './icons.ts'
import { ICON_HEX, ICON_SIZES_CM, loadIconIndex, resolveIcon } from './icons.ts'
import { COLOURS } from './tokens.ts'

export type LibraryId = 'lucide' | 'sectors' | 'tabler' | 'flags' | 'maps'
/** Stroke icons are recoloured through their stroke, maps through their fill; flags keep their own colours. */
export type EntryKind = 'stroke' | 'flag' | 'map'

export interface LibraryInfo {
  id: LibraryId
  label: string
  kind: EntryKind
  licence: string
}

export const LIBRARIES: readonly LibraryInfo[] = [
  { id: 'lucide', label: 'Lucide', kind: 'stroke', licence: 'Lucide icons, ISC licence.' },
  { id: 'sectors', label: 'BwB sectors', kind: 'stroke', licence: 'The 13 sectors on bwb.earth, drawn with Lucide icons (ISC licence).' },
  { id: 'tabler', label: 'Tabler', kind: 'stroke', licence: 'Tabler icons, MIT licence. Brand logos belong to their owners.' },
  { id: 'flags', label: 'Flags', kind: 'flag', licence: 'Flags from flag-icons, MIT licence.' },
  { id: 'maps', label: 'Maps', kind: 'map', licence: 'Maps from Natural Earth via world-atlas, ISC licence.' },
]

interface SearchFields {
  /** The name in lower-case kebab form, and its words. */
  name: string
  parts: string[]
  tags: string[]
  /** Whole queries that count as the name itself: a flag's code and aliases. */
  exact: string[]
}

interface BaseEntry {
  /** Unique across libraries, such as "tabler:droplet" or "flags:ke". */
  key: string
  library: LibraryId
  /** The icon's name, or the flag's or map's code. */
  id: string
  /** What people read: the icon's name, the sector or the country. */
  name: string
  /** Tabler's category, or the region of a flag or map. */
  group: string
  search: SearchFields
}

export interface StrokeEntry extends BaseEntry {
  kind: 'stroke'
  nodes: IconNode[]
}
export interface FlagEntry extends BaseEntry {
  kind: 'flag'
}
export interface MapEntry extends BaseEntry {
  kind: 'map'
  /** Path of the SVG from the add-in's root, and its drawn size, which gives the aspect. */
  file: string
  width: number
  height: number
}
export type IconEntry = StrokeEntry | FlagEntry | MapEntry

export interface Library {
  info: LibraryInfo
  entries: IconEntry[]
  /** BwB topics ("climate finance") and the entry ids that suit them, best first. Lucide only. */
  topics: Record<string, string[]>
}

export interface TablerJson {
  version: string
  licence: string
  icons: Record<string, { n: IconNode[]; t: string[]; c: string }>
}
export interface FlagJson {
  version: string
  licence: string
  flags: Array<{ code: string; name: string; region: string; aliases: string[] }>
}
export interface MapJson {
  version: string
  licence: string
  source: string
  maps: Array<{ code: string; name: string; aliases: string[]; region: string; file: string; w: number; h: number }>
}

/** The 13 sectors on bwb.earth, each with the Lucide icon the website uses. */
export const SECTORS: ReadonlyArray<{ id: string; label: string; icon: string; tags: string[] }> = [
  { id: 'sovereign', label: 'Sovereign', icon: 'landmark', tags: ['government', 'ministry', 'public finance', 'debt'] },
  { id: 'urban', label: 'Urban', icon: 'building-complex', tags: ['city', 'municipal', 'infrastructure'] },
  { id: 'education', label: 'Education', icon: 'book-open', tags: ['school', 'learning', 'knowledge'] },
  { id: 'built-environment', label: 'Built environment', icon: 'building', tags: ['buildings', 'real estate', 'retrofit'] },
  { id: 'energy', label: 'Energy', icon: 'zap', tags: ['power', 'electricity', 'grid', 'renewable'] },
  { id: 'blue-economy', label: 'Blue economy', icon: 'waves-horizontal', tags: ['ocean', 'sea', 'marine', 'water', 'coast'] },
  { id: 'mobility', label: 'Mobility', icon: 'car', tags: ['transport', 'electric vehicle', 'bus'] },
  { id: 'circular-economy', label: 'Circular economy', icon: 'recycle', tags: ['recycling', 'waste', 'materials'] },
  { id: 'biodiversity', label: 'Biodiversity', icon: 'leaf', tags: ['nature', 'nature-based solutions', 'ecosystems'] },
  { id: 'carbon', label: 'Carbon', icon: 'cloud', tags: ['forestry', 'emissions', 'carbon markets'] },
  { id: 'gender', label: 'Gender', icon: 'users', tags: ['inclusion', 'women', 'inclusive finance'] },
  { id: 'food', label: 'Food', icon: 'wheat', tags: ['agriculture', 'farming', 'food security'] },
  { id: 'climate-security', label: 'Climate security', icon: 'shield', tags: ['resilience', 'adaptation', 'drought', 'security'] },
]

/** Lower case without accents or apostrophes, words joined by hyphens: "Côte d’Ivoire" gives "cote-divoire". */
export function fold(text: string): string {
  return text.normalize('NFD').replace(/\p{M}/gu, '').toLowerCase().replace(/['’‘]/g, '').replace(/&/g, ' and ')
    .replace(/[^a-z0-9]+/g, '-').replace(/^-+|-+$/g, '')
}

function fields(name: string, tags: string[], exact: string[] = []): SearchFields {
  const key = fold(name)
  return { name: key, parts: key.split('-'), tags, exact: exact.map(fold) }
}

const STOP_WORDS = new Set(['and', 'the', 'of'])

/** Tags for places: each alias and region as a phrase, and its words. */
function placeTags(phrases: string[]): string[] {
  const folded = phrases.map((p) => fold(p).replace(/-/g, ' ')).filter(Boolean)
  return [...new Set([...folded, ...folded.flatMap((p) => p.split(' ').filter((w) => !STOP_WORDS.has(w)))])]
}

function strokeEntries(library: LibraryId, icons: Record<string, { n: IconNode[]; t: string[]; c?: string }>): StrokeEntry[] {
  return Object.entries(icons).map(([id, icon]) => ({
    key: `${library}:${id}`, library, id, name: id, group: icon.c ?? '', kind: 'stroke', nodes: icon.n, search: fields(id, icon.t),
  }))
}

const infoOf = (id: LibraryId): LibraryInfo => LIBRARIES.find((l) => l.id === id) ?? { id, label: id, kind: 'stroke', licence: '' }

export function lucideLibrary(index: IconIndex): Library {
  const topics: Record<string, string[]> = {}
  for (const [topic, names] of Object.entries(index.topics)) {
    topics[topic] = names.map((n) => resolveIcon(index, n)).filter((n): n is string => n !== null)
  }
  return { info: infoOf('lucide'), entries: strokeEntries('lucide', index.icons), topics }
}

export function sectorLibrary(index: IconIndex): Library {
  const entries: StrokeEntry[] = []
  for (const sector of SECTORS) {
    const icon = index.icons[resolveIcon(index, sector.icon) ?? '']
    if (!icon) continue
    entries.push({
      key: `sectors:${sector.id}`, library: 'sectors', id: sector.id, name: sector.label, group: `Lucide ${sector.icon}`, kind: 'stroke',
      nodes: icon.n, search: fields(sector.label, [...sector.tags, sector.icon, ...icon.t]),
    })
  }
  return { info: infoOf('sectors'), entries, topics: {} }
}

export function tablerLibrary(json: TablerJson): Library {
  return { info: infoOf('tabler'), entries: strokeEntries('tabler', json.icons), topics: {} }
}

export function flagLibrary(json: FlagJson): Library {
  const entries: FlagEntry[] = json.flags.map((f) => ({
    key: `flags:${f.code}`, library: 'flags', id: f.code, name: f.name, group: f.region, kind: 'flag',
    search: fields(f.name, placeTags([...f.aliases, f.region]), [f.code, f.name, ...f.aliases]),
  }))
  return { info: infoOf('flags'), entries, topics: {} }
}

export function mapLibrary(json: MapJson): Library {
  const entries: MapEntry[] = json.maps.map((m) => ({
    key: `maps:${m.code}`, library: 'maps', id: m.code, name: m.name, group: m.region, kind: 'map', file: m.file, width: m.w, height: m.h,
    search: fields(m.name, placeTags([...m.aliases, m.region]), [m.code, m.name, ...m.aliases]),
  }))
  return { info: infoOf('maps'), entries, topics: {} }
}

async function fetchJson<T>(url: string, what: string): Promise<T> {
  const response = await fetch(url)
  if (!response.ok) throw new Error(`The ${what} could not be loaded (${response.status}).`)
  return (await response.json()) as T
}

const loading = new Map<LibraryId, Promise<Library>>()

/** Loads a library once; a failed load is forgotten so that it can be tried again. */
export function loadLibrary(id: LibraryId): Promise<Library> {
  let pending = loading.get(id)
  if (!pending) {
    pending = (async (): Promise<Library> => {
      if (id === 'lucide') return lucideLibrary(await loadIconIndex())
      if (id === 'sectors') return sectorLibrary(await loadIconIndex())
      if (id === 'tabler') return tablerLibrary(await fetchJson<TablerJson>('./assets/icons/tabler.json', 'Tabler icon library'))
      if (id === 'flags') return flagLibrary(await fetchJson<FlagJson>('./assets/icons/flags/flags.json', 'flag library'))
      return mapLibrary(await fetchJson<MapJson>('./assets/icons/maps/maps.json', 'map library'))
    })()
    pending.catch(() => loading.delete(id))
    loading.set(id, pending)
  }
  return pending
}

// ---- Search ----

function queryWords(query: string): { q: string; words: string[] } {
  const q = fold(query).replace(/^lucide-/, '').replace(/-icon$/, '').replace(/-/g, ' ')
  return { q, words: q.split(' ').filter(Boolean) }
}

/** Topic scores for a query, by entry id: 70 for the topic's first icon, a little less for each one after. */
function topicScores(topics: Record<string, string[]>, q: string, words: string[]): Map<string, number> {
  const scores = new Map<string, number>()
  for (const [key, ids] of Object.entries(topics)) {
    const k = key.toLowerCase()
    const parts = k.split(' ')
    if (!(k === q || words.includes(k) || (parts.length > 1 && parts.every((p) => words.includes(p))))) continue
    ids.forEach((id, pos) => scores.set(id, Math.max(scores.get(id) ?? 0, 70 - pos * 0.6)))
  }
  return scores
}

function scoreEntry({ search }: IconEntry, q: string, words: string[], topic: number | undefined): number {
  const { name, parts, tags } = search
  let matched = 0
  let score = 0
  for (const w of words) {
    let s = 0
    if (name === w) s = 100
    else if (parts.includes(w)) s = 60
    else if (name.startsWith(w)) s = 45
    else if (w.length > 2 && name.includes(w)) s = 30
    else if (tags.includes(w)) s = 40
    else if (w.length > 3 && tags.some((t) => t.startsWith(w) || t.includes(w))) s = 20
    if (s) {
      matched += 1
      score += s
    }
  }
  if (search.exact.includes(q.replace(/ /g, '-'))) {
    matched = Math.max(matched, 1)
    score += 100
  }
  if (topic !== undefined) {
    matched = Math.max(matched, 1) + (matched ? 1 : 0)
    score += topic
  }
  return matched ? matched * 1000 + score - name.length * 0.01 : 0
}

/** Ranks entries across libraries: whole and partial name matches, tags, then BwB topics. Ties keep library order. */
export function searchLibraries(libraries: readonly Library[], query: string, limit = 300): IconEntry[] {
  const { q, words } = queryWords(query)
  if (!words.length) return []
  const order = (id: LibraryId): number => LIBRARIES.findIndex((l) => l.id === id)
  const hits: Array<{ entry: IconEntry; score: number }> = []
  for (const library of libraries) {
    const topics = topicScores(library.topics, q, words)
    for (const entry of library.entries) {
      const score = scoreEntry(entry, q, words, topics.get(entry.id))
      if (score) hits.push({ entry, score })
    }
  }
  hits.sort((a, b) => b.score - a.score || order(a.entry.library) - order(b.entry.library) || (a.entry.name < b.entry.name ? -1 : 1))
  return hits.slice(0, limit).map((h) => h.entry)
}

// ---- Colours, sizes and SVG ----

export type LibraryColour = 'green' | 'navy' | 'charcoal' | 'lightGrey' | 'white' | 'lime'
export const LIBRARY_HEX: Readonly<Record<LibraryColour, string>> = { ...ICON_HEX, lightGrey: COLOURS.lightGrey }

/** Green, navy and charcoal on white (maps also light grey, for a backdrop); white or lime on a navy panel; flags none. */
export function coloursFor(kind: EntryKind, onNavy: boolean): LibraryColour[] {
  if (kind === 'flag') return []
  if (onNavy) return ['white', 'lime']
  return kind === 'map' ? ['green', 'navy', 'charcoal', 'lightGrey'] : ['green', 'navy', 'charcoal']
}

export type FlagShape = 'rect' | 'square' | 'round'

export interface SizeOption {
  id: string
  label: string
  pt: number
}

const PT_PER_CM = 72 / 2.54
const cm = (id: string, label: string, value: number): SizeOption => ({ id, label: `${label}, ${value} cm`, pt: value * PT_PER_CM })

/**
 * Sizes for the host. Word has the four BwB icon sizes, and widths for flags and maps (a map's longer side);
 * PowerPoint has small, medium and large in points, and the browser adds a custom size.
 */
export function sizesFor(kind: EntryKind, host: 'word' | 'powerpoint'): SizeOption[] {
  if (host === 'powerpoint') {
    return [{ id: 'small', label: 'Small, 24 pt', pt: 24 }, { id: 'medium', label: 'Medium, 48 pt', pt: 48 }, { id: 'large', label: 'Large, 96 pt', pt: 96 }]
  }
  if (kind === 'flag') return [cm('small', 'Small', 0.8), cm('medium', 'Medium', 2), cm('large', 'Large', 4)]
  if (kind === 'map') return [cm('small', 'Small', 4), cm('medium', 'Medium', 8), cm('large', 'Large', 16)]
  const { inline, callout, marker, panel } = ICON_SIZES_CM
  return [cm('inline', 'In text', inline), cm('callout', 'Callout', callout), cm('marker', 'Marker', marker), cm('panel', 'Panel', panel)]
}

/** The inserted size in points: icons are square, flags 4:3 (or square), maps keep their aspect with `pt` the longer side. */
export function boxFor(entry: IconEntry, pt: number, shape: FlagShape = 'rect'): { widthPt: number; heightPt: number } {
  if (entry.kind === 'flag') return { widthPt: pt, heightPt: shape === 'rect' ? (pt * 3) / 4 : pt }
  if (entry.kind === 'map') {
    const ratio = entry.width / entry.height
    return ratio >= 1 ? { widthPt: pt, heightPt: pt / ratio } : { widthPt: pt * ratio, heightPt: pt }
  }
  return { widthPt: pt, heightPt: pt }
}

function escapeAttr(value: string): string {
  return value.replace(/&/g, '&amp;').replace(/"/g, '&quot;').replace(/</g, '&lt;')
}

/** A stroke icon (Lucide or Tabler) as standalone SVG in one BwB colour, including any filled dots it has. */
export function strokeSvg(nodes: readonly IconNode[], colour: LibraryColour, px = 24): string {
  const hex = `#${LIBRARY_HEX[colour]}`
  const body = nodes.map(([tag, attrs]) => {
    const list = Object.entries(attrs).map(([k, v]) => `${k}="${escapeAttr(v === 'currentColor' ? hex : v)}"`)
    return `<${tag} ${list.join(' ')}/>`
  })
  return `<svg xmlns="http://www.w3.org/2000/svg" width="${px}" height="${px}" viewBox="0 0 24 24" fill="none" stroke="${hex}" ` +
    `stroke-width="2" stroke-linecap="round" stroke-linejoin="round">${body.join('')}</svg>`
}

/** Sets the width and height on an SVG's root element. */
function sized(svg: string, width: number, height: number): string {
  return svg.replace(/<svg\b([^>]*)>/, (_m, attrs: string) => `<svg${attrs.replace(/\s(width|height)="[^"]*"/g, '')} width="${Math.round(width)}" height="${Math.round(height)}">`)
}

export function flagUrl(code: string, shape: FlagShape = 'rect'): string {
  return `./assets/icons/flags/${shape === 'rect' ? '4x3' : '1x1'}/${code}.svg`
}

/** A flag in its own colours: 4:3, square, or square clipped to a circle. */
export function flagSvg(text: string, shape: FlagShape, px: number): string {
  if (shape === 'rect') return sized(text, px, (px * 3) / 4)
  const square = sized(text, px, px)
  if (shape === 'square') return square
  return square
    .replace(/<svg\b[^>]*>/, (open) => `${open}<defs><clipPath id="bwb-round"><circle cx="256" cy="256" r="256"/></clipPath></defs><g clip-path="url(#bwb-round)">`)
    .replace(/<\/svg>\s*$/, '</g></svg>')
}

/** A map silhouette filled in one BwB colour, `px` along its longer side. */
export function mapSvg(text: string, colour: LibraryColour, px: number, entry: Pick<MapEntry, 'width' | 'height'>): string {
  const ratio = entry.width / entry.height
  return sized(text.replace(/currentColor/g, `#${LIBRARY_HEX[colour]}`), ratio >= 1 ? px : px * ratio, ratio >= 1 ? px / ratio : px)
}

const texts = new Map<string, Promise<string>>()

function fetchText(url: string, what: string): Promise<string> {
  let pending = texts.get(url)
  if (!pending) {
    pending = fetch(url).then((r) => {
      if (!r.ok) throw new Error(`The ${what} could not be loaded (${r.status}).`)
      return r.text()
    })
    pending.catch(() => texts.delete(url))
    texts.set(url, pending)
  }
  return pending
}

/** Standalone SVG for any entry, `px` wide (a map's longer side). Flags ignore the colour. */
export async function iconSvgFor(entry: IconEntry, colour: LibraryColour, px: number, shape: FlagShape = 'rect'): Promise<string> {
  if (entry.kind === 'stroke') return strokeSvg(entry.nodes, colour, px)
  if (entry.kind === 'flag') return flagSvg(await fetchText(flagUrl(entry.id, shape), 'flag'), shape, px)
  return mapSvg(await fetchText(`./${entry.file}`, 'map'), colour, px, entry)
}

/** Alt text in the BwB pattern: "Icon: droplet", "Flag: Kenya", "Map: Kenya". */
export function altTextFor(entry: IconEntry): string {
  if (entry.kind === 'flag') return `Flag: ${entry.name}`
  if (entry.kind === 'map') return `Map: ${entry.name}`
  return `Icon: ${entry.library === 'sectors' ? entry.name : entry.name.replace(/-/g, ' ')}`
}

/** What "Copy name" copies: the icon name the BwB skill's {icon:name} token takes, or the place's name. */
export function copyTextFor(entry: IconEntry): string {
  if (entry.library === 'sectors') return SECTORS.find((s) => s.id === entry.id)?.icon ?? entry.name
  return entry.name
}
