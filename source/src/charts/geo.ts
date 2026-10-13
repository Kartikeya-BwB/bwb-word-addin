// The world map behind data maps (public/assets/charts/world.json, made by scripts/make-chart-maps.mjs): types,
// a check that a loaded file is one, country name matching, and the geometry the renderer draws from. Pure: the
// builder loads the file once and passes it in.

export type RegionId = 'world' | 'africa' | 'europe' | 'asia' | 'latin-america' | 'mena' | 'sub-saharan-africa'

export const REGION_IDS: readonly RegionId[] = ['world', 'africa', 'europe', 'asia', 'latin-america', 'mena', 'sub-saharan-africa']

export interface WorldCountry {
  /** ISO 3166-1 alpha-2 and alpha-3 codes. */
  a2: string
  a3: string
  /** Name, as BwB writes it. */
  n: string
  /** Other names people paste. */
  al: string[]
  /** World Bank region and continent. */
  r: string
  c: string
  /** Outline in the projected plane, or empty when the country is too small to draw. */
  d: string
  /** Bounds of the outline in the plane. */
  b: [number, number, number, number] | null
}

export interface WorldRegion {
  id: RegionId
  name: string
  /** The frame to fit in the chart, in the projected plane. */
  box: [number, number, number, number]
  /** Alpha-2 codes of the countries in the region. */
  members: string[]
}

export interface WorldMap {
  version: string
  source: string
  credit: string
  projection: string
  width: number
  height: number
  countries: WorldCountry[]
  regions: WorldRegion[]
}

const isBox = (v: unknown): v is [number, number, number, number] => Array.isArray(v) && v.length === 4 && v.every((n) => typeof n === 'number')

/** True when a loaded file has the shape the renderer needs. */
export function isWorldMap(value: unknown): value is WorldMap {
  if (!value || typeof value !== 'object') return false
  const w = value as Partial<WorldMap>
  return (
    typeof w.width === 'number' && typeof w.height === 'number' && Array.isArray(w.countries) && Array.isArray(w.regions) &&
    w.countries.every((c) => typeof c?.a2 === 'string' && typeof c.n === 'string' && typeof c.d === 'string' && Array.isArray(c.al)) &&
    w.regions.every((r) => typeof r?.id === 'string' && isBox(r.box) && Array.isArray(r.members))
  )
}

/** A name reduced to what matters for matching: no accents, case, punctuation, "the" or "St." versus "Saint". */
export function normaliseName(text: string): string {
  return text
    .normalize('NFD')
    .replace(/[\u{300}-\u{36F}]/gu, '')
    .toLowerCase()
    .replace(/&/g, ' and ')
    .replace(/['\u{2018}\u{2019}`]/gu, '')
    .replace(/\bst\b\.?/g, 'saint')
    .replace(/[^a-z0-9]+/g, ' ')
    .trim()
    .replace(/^the /, '')
    .replace(/ the$/, '')
}

const matchers = new WeakMap<WorldMap, Map<string, WorldCountry>>()

function indexOf(world: WorldMap): Map<string, WorldCountry> {
  const cached = matchers.get(world)
  if (cached) return cached
  const index = new Map<string, WorldCountry>()
  for (const c of world.countries) {
    for (const name of [c.n, ...c.al]) {
      const key = normaliseName(name)
      if (key && !index.has(key)) index.set(key, c)
    }
  }
  for (const c of world.countries) {
    index.set(`#${c.a2.toLowerCase()}`, c)
    index.set(`#${c.a3.toLowerCase()}`, c)
  }
  matchers.set(world, index)
  return index
}

/** The country a cell names, by name, other name, or ISO alpha-2 or alpha-3 code; null when none matches. */
export function matchCountry(world: WorldMap, text: string): WorldCountry | null {
  const index = indexOf(world)
  const trimmed = text.trim()
  if (/^[A-Za-z]{2,3}$/.test(trimmed)) {
    const byCode = index.get(`#${trimmed.toLowerCase()}`)
    if (byCode) return byCode
  }
  return index.get(normaliseName(trimmed)) ?? null
}

export const regionOf = (world: WorldMap, id: RegionId): WorldRegion | undefined => world.regions.find((r) => r.id === id)

/** A ring as flat coordinates: x0, y0, x1, y1 and so on. */
export type Ring = number[]

const ringCache = new WeakMap<WorldCountry, Ring[]>()

/** The rings of a country's outline, read from its path once and kept. */
export function countryRings(country: WorldCountry): Ring[] {
  const cached = ringCache.get(country)
  if (cached) return cached
  const rings: Ring[] = []
  for (const part of country.d.split('M')) {
    const numbers = part.match(/-?\d+(?:\.\d+)?/g)?.map(Number) ?? []
    if (numbers.length < 6) continue
    let x = numbers[0] ?? 0
    let y = numbers[1] ?? 0
    const ring: Ring = [x, y]
    for (let i = 2; i + 1 < numbers.length; i += 2) {
      x += numbers[i] ?? 0
      y += numbers[i + 1] ?? 0
      ring.push(x, y)
    }
    rings.push(ring)
  }
  ringCache.set(country, rings)
  return rings
}

type Box = [number, number, number, number]

/** True when two boxes overlap. */
export const boxesMeet = (a: Box, b: Box): boolean => a[0] <= b[2] && b[0] <= a[2] && a[1] <= b[3] && b[1] <= a[3]

/**
 * Clips a ring to a box, one edge at a time (Sutherland and Hodgman's method). A concave shape can come out with
 * edges running along the box, which a filled outline hides.
 */
export function clipRing(ring: Ring, [x0, y0, x1, y1]: Box): Ring {
  const edges: Array<{ inside: (x: number, y: number) => boolean; cut: (ax: number, ay: number, bx: number, by: number) => [number, number] }> = [
    { inside: (x) => x >= x0, cut: (ax, ay, bx, by) => [x0, ay + ((by - ay) * (x0 - ax)) / (bx - ax)] },
    { inside: (x) => x <= x1, cut: (ax, ay, bx, by) => [x1, ay + ((by - ay) * (x1 - ax)) / (bx - ax)] },
    { inside: (_x, y) => y >= y0, cut: (ax, ay, bx, by) => [ax + ((bx - ax) * (y0 - ay)) / (by - ay), y0] },
    { inside: (_x, y) => y <= y1, cut: (ax, ay, bx, by) => [ax + ((bx - ax) * (y1 - ay)) / (by - ay), y1] },
  ]
  let points = ring
  for (const edge of edges) {
    const out: Ring = []
    const n = points.length / 2
    for (let i = 0; i < n; i += 1) {
      const ax = points[(2 * ((i + n - 1) % n))] ?? 0
      const ay = points[(2 * ((i + n - 1) % n)) + 1] ?? 0
      const bx = points[2 * i] ?? 0
      const by = points[2 * i + 1] ?? 0
      const aIn = edge.inside(ax, ay)
      const bIn = edge.inside(bx, by)
      if (bIn) {
        if (!aIn) out.push(...edge.cut(ax, ay, bx, by))
        out.push(bx, by)
      } else if (aIn) out.push(...edge.cut(ax, ay, bx, by))
    }
    points = out
    if (points.length < 6) return []
  }
  return points
}
