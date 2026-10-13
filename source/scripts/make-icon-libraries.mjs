// Builds the icon libraries the task pane's icon browser offers beside Lucide (scripts/make-icon-index.mjs):
//   public/assets/icons/tabler.json      Tabler outline icons: drawing nodes, tags and category (MIT)
//   public/assets/icons/flags/           flag-icons SVGs, 4x3 and 1x1, and flags.json (MIT)
//   public/assets/icons/maps/            country and region silhouettes drawn from Natural Earth via world-atlas (ISC)
// The output is committed; run `node scripts/make-icon-libraries.mjs` again after updating those packages.
import { copyFileSync, mkdirSync, readFileSync, readdirSync, rmSync, statSync, writeFileSync } from 'node:fs'
import { createRequire } from 'node:module'
import { dirname, join } from 'node:path'
import { fileURLToPath } from 'node:url'
import {
  geoArea, geoAzimuthalEqualArea, geoBounds, geoCentroid, geoContains, geoEqualEarth, geoIdentity, geoPath, geoStream,
} from 'd3-geo'

const require = createRequire(import.meta.url)
const { feature, merge } = require('topojson-client')

const root = join(dirname(fileURLToPath(import.meta.url)), '..')
const modules = join(root, 'node_modules')
const out = join(root, 'public', 'assets', 'icons')
const readJson = (path) => JSON.parse(readFileSync(path, 'utf8'))
const versionOf = (pkg) => readJson(join(modules, pkg, 'package.json')).version
const trimmed = []

rmSync(out, { recursive: true, force: true })
mkdirSync(out, { recursive: true })

// ---- Tabler: the same compact shape as the Lucide index ----
const tablerDir = join(modules, '@tabler', 'icons')
const [tablerMeta, tablerNodes] = ['icons.json', 'tabler-nodes-outline.json'].map((f) => readJson(join(tablerDir, f)))
const tabler = {}
for (const name of Object.keys(tablerNodes).sort()) {
  const meta = tablerMeta[name] ?? {}
  const words = new Set(name.split('-'))
  // Tags that repeat the name add nothing to search: the name's own words already match.
  const tags = [...new Set((meta.tags ?? []).map((t) => String(t).toLowerCase().trim()))].filter((t) => t && t !== name && !words.has(t))
  tabler[name] = { n: tablerNodes[name], t: tags, c: meta.category ?? '' }
}
writeFileSync(join(out, 'tabler.json'), JSON.stringify({ version: versionOf('@tabler/icons'), licence: 'MIT', icons: tabler }))
copyFileSync(join(tablerDir, 'LICENSE'), join(out, 'LICENSE-tabler.txt'))

// ---- Names and regions shared by flags and maps ----
const longName = new Intl.DisplayNames('en-GB', { type: 'region' })
const shortName = new Intl.DisplayNames('en-GB', { type: 'region', style: 'short' })
// CLDR's names for these read badly in a document or in alt text.
const NAME_FIXES = { CD: 'Democratic Republic of the Congo', CG: 'Republic of the Congo', MM: 'Myanmar', HK: 'Hong Kong', MO: 'Macao' }
const nameOf = (code) => (NAME_FIXES[code.toUpperCase()] ?? longName.of(code.toUpperCase())).replace(/ & /g, ' and ')

// World Bank regions, with territories placed by geography.
const REGIONS = {
  'East Asia and Pacific': 'AS AU BN CC CK CN CX FJ FM GU HK ID JP KH KI KP KR LA MH MM MN MO MP MY NC NF NR NU NZ PF PG PH PN PW SB SG TH TK TL TO TV TW UM VN VU WF WS',
  'Europe and Central Asia': 'AD AL AM AT AX AZ BA BE BG BY CH CY CZ DE DK EE ES FI FO FR GB GE GG GI GL GR HR HU IE IM IS IT JE KG KZ LI LT LU LV MC MD ME MK NL NO PL PT RO RS RU SE SI SJ SK SM TJ TM TR UA UZ VA XK',
  'Latin America and the Caribbean': 'AG AI AR AW BB BL BO BQ BR BS BZ CL CO CR CU CW DM DO EC FK GD GF GP GS GT GY HN HT JM KN KY LC MF MQ MS MX NI PA PE PR PY SR SV SX TC TT UY VC VE VG VI',
  'Middle East and North Africa': 'AE BH DJ DZ EG EH IL IQ IR JO KW LB LY MA MT OM PS QA SA SY TN YE',
  'North America': 'BM CA PM US',
  'South Asia': 'AF BD BT IN IO LK MV NP PK',
  'Sub-Saharan Africa': 'AO BF BI BJ BW CD CF CG CI CM CV ER ET GA GH GM GN GQ GW KE KM LR LS MG ML MR MU MW MZ NA NE NG RE RW SC SD SH SL SN SO SS ST SZ TD TG TZ UG YT ZA ZM ZW',
  Antarctica: 'AQ BV HM TF',
}
const regionByCode = new Map(Object.entries(REGIONS).flatMap(([region, codes]) => codes.split(' ').map((c) => [c, region])))
function regionOf(code) {
  const region = regionByCode.get(code.toUpperCase())
  if (region) return region
  throw new Error(`No region for ${code}: add it to REGIONS.`)
}

// ---- Flags ----
const flagDir = join(modules, 'flag-icons')
const countries = readJson(join(flagDir, 'country.json')).sort((a, b) => a.code.localeCompare(b.code))
const continentOf = new Map(countries.map((c) => [c.code.toUpperCase(), c.continent]))
// Left out: Spain's autonomous communities, Northern Ireland (no official flag since 1972), the Canary Islands,
// Clipperton and Diego Garcia (they fly the French and BIOT flags), Ascension, Tristan da Cunha and the "unknown" flag.
const SKIP_FLAGS = new Set(['es-ct', 'es-ga', 'es-pv', 'gb-nir', 'ic', 'cp', 'dg', 'sh-ac', 'sh-ta', 'xx'])
// flag-icons draws "sh" as the Union Flag; St Helena's own flag is "sh-hl".
const FLAG_SOURCE = { sh: 'sh-hl' }
const ORGANISATIONS = new Set(['eu', 'un', 'arab', 'asean', 'cefta', 'eac', 'pc'])
const FLAG_ALIASES = {
  ae: 'UAE, Emirates', ba: 'Bosnia', bs: 'The Bahamas', cd: 'DRC, DR Congo, Congo-Kinshasa', cg: 'Congo-Brazzaville, Congo',
  ci: 'Ivory Coast', cv: 'Cabo Verde', cz: 'Czech Republic', fm: 'Federated States of Micronesia', gb: 'UK, Britain, Great Britain',
  gm: 'The Gambia', kp: 'DPRK, Korea', kr: 'Korea, Republic of Korea', la: 'Lao PDR', mk: 'Macedonia', mm: 'Burma',
  nl: 'Holland', ps: 'Palestine, West Bank, Gaza', ru: 'Russian Federation', sz: 'Swaziland', tl: 'East Timor', tr: 'Turkey',
  us: 'USA, America, United States of America', va: 'Holy See, Vatican', vn: 'Viet Nam', sy: 'Syrian Arab Republic',
  eu: 'EU', un: 'UN', 'gb-eng': 'English', 'gb-sct': 'Scottish', 'gb-wls': 'Welsh, Cymru', arab: 'Arab League',
  asean: 'ASEAN', eac: 'EAC', pc: 'SPC, Secretariat of the Pacific Community', cefta: 'CEFTA',
}
const minifySvg = (text) => text.replace(/>\s+</g, '><').trim()
const flagNames = new Map(countries.map((c) => [c.code, c.name]))

/** The name and search aliases for a lower-case code. Codes CLDR knows take its name; "pc" is the Pacific Community. */
function namesFor(code) {
  const cldr = /^[a-z]{2}$/.test(code) && longName.of(code.toUpperCase()) !== code.toUpperCase()
  const extra = (FLAG_ALIASES[code] ?? '').split(', ')
  const name = cldr ? nameOf(code) : flagNames.get(code) ?? code
  const candidates = cldr ? [flagNames.get(code) ?? '', longName.of(code.toUpperCase()), shortName.of(code.toUpperCase()), ...extra] : extra
  const aliases = [...new Set(candidates.map((a) => a.replace(/ & /g, ' and ').replace(/ - /g, '-')))]
    .filter((a) => a && a !== name && a.toUpperCase() !== code.toUpperCase())
  return { name, aliases }
}

const flags = []
for (const dir of ['4x3', '1x1']) mkdirSync(join(out, 'flags', dir), { recursive: true })
for (const { code } of countries) {
  if (Object.values(FLAG_SOURCE).includes(code)) continue
  if (SKIP_FLAGS.has(code)) {
    trimmed.push(`flag ${code}`)
    continue
  }
  const source = (dir) => readFileSync(join(flagDir, 'flags', dir, `${FLAG_SOURCE[code] ?? code}.svg`), 'utf8')
  for (const dir of ['4x3', '1x1']) writeFileSync(join(out, 'flags', dir, `${code}.svg`), minifySvg(source(dir)))
  const { name, aliases } = namesFor(code)
  const region = ORGANISATIONS.has(code) ? 'International organisations' : regionOf(code.slice(0, 2))
  flags.push({ code, name, region, aliases })
}
writeFileSync(join(out, 'flags', 'flags.json'), JSON.stringify({ version: versionOf('flag-icons'), licence: 'MIT', flags }))
copyFileSync(join(flagDir, 'LICENSE'), join(out, 'flags', 'LICENSE.txt'))

// ---- Geometry helpers ----
const [SIZE, MARGIN] = [1000, 4]

/** Douglas-Peucker on an open line; keeps both ends. */
function simplify(points, tolerance) {
  if (points.length < 3 || tolerance <= 0) return points
  const keep = new Uint8Array(points.length)
  keep[0] = keep[points.length - 1] = 1
  const stack = [[0, points.length - 1]]
  while (stack.length) {
    const [a, b] = stack.pop()
    const [[ax, ay], [bx, by]] = [points[a], points[b]]
    const length = Math.hypot(bx - ax, by - ay)
    let [worst, far] = [-1, tolerance]
    for (let i = a + 1; i < b; i++) {
      const [px, py] = points[i]
      const d = length ? Math.abs((by - ay) * px - (bx - ax) * py + bx * ay - by * ax) / length : Math.hypot(px - ax, py - ay)
      if (d > far) [far, worst] = [d, i]
    }
    if (worst > 0) {
      keep[worst] = 1
      stack.push([a, worst], [worst, b])
    }
  }
  return points.filter((_, i) => keep[i])
}

/** Simplifies a closed ring (without its closing point), split at its farthest point so both halves keep their shape. */
function simplifyRing(ring, tolerance) {
  if (ring.length < 4 || tolerance <= 0) return ring
  const [x0, y0] = ring[0]
  const distances = ring.map(([x, y]) => Math.hypot(x - x0, y - y0))
  const split = distances.reduce((best, d, i) => (d > distances[best] ? i : best), 0)
  return [...simplify(ring.slice(0, split + 1), tolerance), ...simplify([...ring.slice(split), ring[0]], tolerance).slice(1, -1)]
}

/**
 * Simplifies the topology's shared arcs (in degrees), so that neighbours keep identical borders after merging and
 * cutting. Arcs are split where they cross the antimeridian.
 */
function simplifyTopology(topology, tolerance) {
  const [sx, sy] = topology.transform.scale
  const arcs = topology.arcs.map((arc) => {
    let [x, y] = [0, 0]
    const points = arc.map(([dx, dy]) => [(x += dx) * sx, (y += dy) * sy])
    const last = points.length - 1
    let kept = []
    if (last > 3 && points[0][0] === points[last][0] && points[0][1] === points[last][1]) {
      const ring = simplifyRing(points.slice(0, last), tolerance)
      kept = ring.length >= 3 ? [...ring, ring[0]] : points
    } else {
      let start = 0
      for (let i = 1; i <= points.length; i++) {
        if (i < points.length && Math.abs(points[i][0] - points[i - 1][0]) <= 180) continue
        kept.push(...simplify(points.slice(start, i), tolerance))
        start = i
      }
    }
    const quantised = kept.map(([kx, ky]) => [Math.round(kx / sx), Math.round(ky / sy)])
    return quantised.map(([qx, qy], i) => (i ? [qx - quantised[i - 1][0], qy - quantised[i - 1][1]] : [qx, qy]))
  })
  return { ...topology, arcs }
}

/** Cuts a lon/lat geometry to a box on the plane, after unwrapping longitudes west of `from` by 360 degrees. */
function cutLonLat(geometry, [[x0, y0], [x1, y1]], from = -180) {
  const unwrap = (lon) => (lon < from ? lon + 360 : lon)
  const shifted = JSON.parse(JSON.stringify(geometry), (_, v) => (Array.isArray(v) && typeof v[0] === 'number' ? [unwrap(v[0]), v[1]] : v))
  const polygons = []
  let [rings, ring] = [[], []]
  const sink = {
    point: (x, y) => void ring.push([x, -y]), lineStart: () => void (ring = []), sphere: () => {},
    lineEnd: () => void (ring.length > 2 && rings.push([...ring, ring[0]])),
    polygonStart: () => void (rings = []), polygonEnd: () => void (rings.length && polygons.push(rings)),
  }
  // Mirrored so that the plane has the same orientation as a projected map, which d3's rectangle clip expects.
  geoStream(shifted, geoIdentity().reflectY(true).clipExtent([[x0, -y1], [x1, -y0]]).stream(sink))
  return { type: 'MultiPolygon', coordinates: polygons }
}

const ringArea = (r) => r.reduce((sum, [x, y], i) => sum + x * r[(i + 1) % r.length][1] - r[(i + 1) % r.length][0] * y, 0) / 2

const numbers = (values) => values.map((n, i) => (i && n >= 0 ? ` ${n}` : `${n}`)).join('')

const sphericalArea = (g) => Math.min(geoArea(g), 4 * Math.PI - geoArea(g))

/** Reverses any polygon whose rings wind the wrong way, which d3 would read as the rest of the globe. */
function outward(geometry) {
  const polygons = geometry.type === 'MultiPolygon' ? geometry.coordinates : [geometry.coordinates]
  const fixed = polygons.map((rings) => (geoArea({ type: 'Polygon', coordinates: rings }) > 2 * Math.PI ? rings.map((r) => [...r].reverse()) : rings))
  return { type: 'MultiPolygon', coordinates: fixed }
}

const boxOf = (ring) => ring.reduce(([x0, y0, x1, y1], [x, y]) => [Math.min(x0, x), Math.min(y0, y), Math.max(x1, x), Math.max(y1, y)],
  [Infinity, Infinity, -Infinity, -Infinity])

/**
 * The rings worth drawing: everything but specks smaller than `minShare` of the whole and, where one island
 * dominates, small islands farther from it than its own size (Sombrero beside Anguilla, Rotuma beside Fiji).
 */
function keepRings(rings, minShare) {
  const areas = rings.map((r) => Math.abs(ringArea(r)))
  const total = areas.reduce((a, b) => a + b, 0)
  const main = areas.indexOf(Math.max(...areas))
  const [mx0, my0, mx1, my1] = boxOf(rings[main] ?? [])
  const reach = Math.hypot(mx1 - mx0, my1 - my0)
  return rings.filter((ring, i) => {
    if (i === main) return true
    if (areas[i] < total * minShare) return false
    if (areas[main] < total / 4 || areas[i] >= total / 20) return true
    const [x0, y0, x1, y1] = boxOf(ring)
    return Math.hypot(Math.max(0, mx0 - x1, x0 - mx1), Math.max(0, my0 - y1, y0 - my1)) <= reach
  })
}

/**
 * Projects a shape, keeps the rings that read at icon size, fits them to a SIZE box, simplifies them and writes one
 * filled path with relative coordinates. Returns null when the result would not be a recognisable shape: a scatter of
 * atolls, or an outline too coarse to be the place's real shape.
 */
function drawShape(geometry, { centre, projection, minShare = 0.0005, tolerance = 0.9, cutEastOf }) {
  const shape = outward(geometry)
  const p = projection ?? geoAzimuthalEqualArea().rotate([-centre[0], -centre[1]])
  p.precision(0.2).fitExtent([[0, 0], [SIZE, SIZE]], shape)
  // A straight cut down the right-hand side, as a map frame would make, through the point given in degrees.
  const edge = cutEastOf && p(cutEastOf)
  if (edge) p.clipExtent([[-SIZE * 10, -SIZE * 10], [edge[0], SIZE * 10]])
  const rings = []
  let ring = []
  const context = { moveTo: (x, y) => rings.push((ring = [[x, y]])), lineTo: (x, y) => void ring.push([x, y]), closePath() {}, arc() {} }
  geoPath(p, context)(shape)
  const kept = keepRings(rings, minShare)
  const [x0, y0, x1, y1] = boxOf(kept.flat())
  const scale = (SIZE - 2 * MARGIN) / Math.max(x1 - x0, y1 - y0)
  const [lx0, ly0, lx1, ly1] = boxOf(kept.reduce((a, b) => (Math.abs(ringArea(b)) > Math.abs(ringArea(a)) ? b : a)))
  if (Math.max(lx1 - lx0, ly1 - ly0) * scale < SIZE / 12) return null
  const width = Math.round((x1 - x0) * scale + 2 * MARGIN)
  const height = Math.round((y1 - y0) * scale + 2 * MARGIN)
  const paths = []
  let outline = [0, 0]
  for (const r of kept) {
    const fitted = r.map(([x, y]) => [(x - x0) * scale + MARGIN, (y - y0) * scale + MARGIN])
    const points = simplifyRing(fitted, tolerance).map(([x, y]) => [Math.round(x), Math.round(y)])
      .filter((pt, i, all) => i === 0 || pt[0] !== all[i - 1][0] || pt[1] !== all[i - 1][1])
    const end = points.at(-1)
    if (end && points.length > 1 && end[0] === points[0][0] && end[1] === points[0][1]) points.pop()
    if (points.length < 3 || Math.abs(ringArea(points)) < 1) continue
    if (Math.abs(ringArea(points)) > outline[0]) outline = [Math.abs(ringArea(points)), points.length]
    const deltas = points.slice(1).flatMap(([x, y], i) => [x - points[i][0], y - points[i][1]])
    paths.push(`M${points[0][0]} ${points[0][1]}l${numbers(deltas)}z`)
  }
  // Below a dozen points the data is too coarse to be the place's real shape (Gibraltar, Monaco).
  if (outline[1] < 12) return null
  const svg = `<svg xmlns="http://www.w3.org/2000/svg" viewBox="0 0 ${width} ${height}" width="${width}" height="${height}">` +
    `<path fill="currentColor" d="${paths.join('')}"/></svg>`
  return { svg, width, height }
}

/** The centre of a shape's largest polygon, so that small outlying islands do not pull the projection away. */
function centreOf(geometry) {
  const polygons = geometry.type === 'MultiPolygon' ? geometry.coordinates : [geometry.coordinates]
  const largest = polygons.map((coordinates) => ({ type: 'Polygon', coordinates }))
    .reduce((a, b) => (sphericalArea(b) > sphericalArea(a) ? b : a))
  return geoBounds(geometry)[0][1] < -89 ? [0, -90] : geoCentroid(largest)
}

// ---- Maps: countries ----
const atlasDir = join(modules, 'world-atlas')
const topology = readJson(join(atlasDir, 'countries-50m.json'))
const mapsOut = join(out, 'maps')
mkdirSync(mapsOut, { recursive: true })

// Natural Earth units without an ISO code: Somaliland and Northern Cyprus join Somalia and Cyprus, as the UK and the
// UN recognise them, and India administers the Siachen Glacier. Two specks of Australian territory are left out.
const UNIT_CODES = { Somaliland: 'SO', 'N. Cyprus': 'CY', 'Siachen Glacier': 'IN', Kosovo: 'XK' }
const SKIP_UNITS = new Set(['Ashmore and Cartier Is.', 'Indian Ocean Ter.'])
const units = new Map()
for (const g of topology.objects.countries.geometries) {
  const name = g.properties.name
  const code = SKIP_UNITS.has(name) ? null : UNIT_CODES[name] ?? new Intl.Locale(`und-${g.id}`).region
  if (code) units.set(code, [...(units.get(code) ?? []), g])
  else trimmed.push(`map ${name}`)
}
const unitsOf = (codes) => codes.flatMap((c) => units.get(c) ?? [])

// Natural Earth 4.1 draws Crimea as part of Russia. The UK and the UN recognise it as part of Ukraine.
const russia = units.get('RU')?.[0]
const crimea = russia?.arcs.findIndex((polygon) => geoContains(feature(topology, { type: 'Polygon', arcs: polygon }), [34.1, 45.0])) ?? -1
if (russia && crimea >= 0) units.get('UA')?.push({ type: 'Polygon', arcs: russia.arcs.splice(crimea, 1)[0] })
const shapes = new Map([...units].map(([code, geometries]) => [code, merge(topology, geometries)]))

// Natural Earth gives Morocco the part of Western Sahara west of the berm. UN maps draw the boundary along 27°40′ N.
const saharan = merge(topology, unitsOf(['MA', 'EH']))
shapes.set('MA', cutLonLat(saharan, [[-180, 27.6667], [180, 90]]))
shapes.set('EH', cutLonLat(saharan, [[-180, -90], [180, 27.6667]]))

// Small countries come out as crude polygons at 1:50m, so they are drawn from the 1:10m data, which also has
// Gibraltar, Tuvalu and the US Minor Outlying Islands. Cyprus stays at 1:50m, where its parts are already joined.
const detailed = readJson(join(atlasDir, 'countries-10m.json'))
const fine = new Map()
for (const g of detailed.objects.countries.geometries) {
  const name = g.properties.name
  const code = name === 'Kosovo' ? 'XK' : g.id && !SKIP_UNITS.has(name) ? new Intl.Locale(`und-${g.id}`).region : null
  if (code && code !== 'CY') fine.set(code, [...(fine.get(code) ?? []), g])
}
for (const [code, geometries] of fine) {
  if (!shapes.has(code) || sphericalArea(shapes.get(code)) * 6371 ** 2 < 35000) shapes.set(code, merge(detailed, geometries))
}

// Remote islands and overseas regions shrink the familiar shape to a speck, so these are cut to their main territory.
const BOXES = {
  US: [[-125, 24.3], [-66.8, 49.5]], FR: [[-5.3, 41.3], [9.7, 51.2]], NL: [[3.3, 50.7], [7.3, 53.6]],
  NO: [[4.5, 57.9], [31.2, 71.3]], PT: [[-9.6, 36.9], [-6.1, 42.2]], ES: [[-9.4, 35.9], [4.4, 43.8]],
  EC: [[-81.1, -5.1], [-75.1, 1.5]], CL: [[-75.8, -56], [-66.3, -17.4]], NZ: [[166.3, -47.4], [178.7, -34.3]],
  AU: [[112.8, -43.8], [153.7, -9.1]], ZA: [[16.4, -35], [33, -22.1]], MU: [[57.3, -20.6], [57.9, -19.9]],
  BR: [[-74, -34], [-34.7, 5.3]], CO: [[-79.1, -4.3], [-66.8, 12.5]], SH: [[-5.9, -16.1], [-5.6, -15.8]], SC: [[55.2, -4.9], [56, -4.2]],
}

const maps = []
for (const code of [...shapes.keys()].sort()) {
  const shape = BOXES[code] ? cutLonLat(shapes.get(code), BOXES[code]) : shapes.get(code)
  const drawn = drawShape(shape, { centre: centreOf(outward(shape)) })
  if (!drawn) {
    trimmed.push(`map ${code}`)
    continue
  }
  const file = `${code.toLowerCase()}.svg`
  writeFileSync(join(mapsOut, file), drawn.svg)
  maps.push({ code: code.toLowerCase(), ...namesFor(code.toLowerCase()), region: regionOf(code), file: `assets/icons/maps/${file}`, w: drawn.width, h: drawn.height })
}

// ---- Maps: regions, merged so that no inner border shows ----
const simple = simplifyTopology(topology, 0.03)
const inContinent = (name) => [...units.keys()].filter((c) => continentOf.get(c) === name)
const inRegion = (name) => [...units.keys()].filter((c) => regionByCode.get(c) === name)
const URALS = 60
const REGION_MAPS = [
  { code: 'africa', name: 'Africa', members: inContinent('Africa'), centre: [17, 2], box: [[-26, -36], [64, 38]] },
  { code: 'asia', name: 'Asia', members: [...inContinent('Asia'), 'RU'], cut: { RU: [[URALS, -90], [200, 90]] }, centre: [95, 38], box: [[25, -12], [192, 78]], from: 0 },
  { code: 'europe', name: 'Europe', members: inContinent('Europe'), centre: [15, 53], box: [[-25, 34], [70, 72]], from: -30, cutEastOf: [45, 55] },
  { code: 'latin-america', name: 'Latin America and the Caribbean', aliases: ['LAC', 'Latin America', 'Caribbean'], members: inRegion('Latin America and the Caribbean'), centre: [-75, -10], box: [[-118, -56], [-33, 33]] },
  { code: 'mena', name: 'Middle East and North Africa', aliases: ['MENA', 'Middle East', 'North Africa'], members: inRegion('Middle East and North Africa'), centre: [22, 28], box: [[-18, 11], [64, 42]] },
  { code: 'sub-saharan-africa', name: 'Sub-Saharan Africa', aliases: ['SSA'], members: inRegion('Sub-Saharan Africa'), centre: [20, -6], box: [[-26, -36], [64, 28]] },
  { code: 'south-asia', name: 'South Asia', members: inRegion('South Asia'), centre: [78, 20], box: [[59, -1.5], [98, 39]] },
  { code: 'east-asia-pacific', name: 'East Asia and Pacific', aliases: ['EAP', 'Asia Pacific'], members: inRegion('East Asia and Pacific'), centre: [130, 5], box: [[72, -48], [190, 55]], from: 0 },
]

const mustDraw = (shape, options) => drawShape(shape, options) ?? (() => { throw new Error('A region map came out empty.') })()
const regionMaps = []
for (const region of REGION_MAPS) {
  const cut = region.cut ?? {}
  const polygons = [...merge(simple, unitsOf(region.members.filter((c) => !cut[c]))).coordinates]
  for (const [code, box] of Object.entries(cut)) polygons.push(...cutLonLat(merge(simple, unitsOf([code])), box, 0).coordinates)
  const shape = cutLonLat({ type: 'MultiPolygon', coordinates: polygons }, region.box, region.from)
  const { svg, width, height } = mustDraw(shape, { centre: region.centre, minShare: 0.00002, tolerance: 0, cutEastOf: region.cutEastOf })
  writeFileSync(join(mapsOut, `${region.code}.svg`), svg)
  regionMaps.push({ code: region.code, name: region.name, aliases: region.aliases ?? [], region: 'Regions', file: `assets/icons/maps/${region.code}.svg`, w: width, h: height })
}

// The world on Equal Earth, from the 1:110m land outline (finer detail does not show at page width), without Antarctica.
const land = readJson(join(atlasDir, 'land-110m.json'))
const landShape = merge(land, land.objects.land.geometries)
const inhabited = { type: 'MultiPolygon', coordinates: landShape.coordinates.filter((polygon) => polygon[0].some(([, lat]) => lat > -60)) }
const world = mustDraw(inhabited, { projection: geoEqualEarth(), minShare: 0.00002 })
writeFileSync(join(mapsOut, 'world.svg'), world.svg)
regionMaps.unshift({ code: 'world', name: 'World', aliases: ['Global', 'Globe'], region: 'Regions', file: 'assets/icons/maps/world.svg', w: world.width, h: world.height })

writeFileSync(join(mapsOut, 'maps.json'), JSON.stringify({
  version: versionOf('world-atlas'), licence: 'ISC', source: 'Natural Earth 4.1.0 via world-atlas', maps: [...regionMaps, ...maps],
}))
writeFileSync(join(mapsOut, 'LICENSE.txt'),
  `${readFileSync(join(atlasDir, 'LICENSE'), 'utf8')}\nThe map data comes from Natural Earth (naturalearthdata.com), which is in the public domain.\n`)

// ---- Totals ----
const folderSize = (dir) => readdirSync(dir, { withFileTypes: true })
  .reduce((sum, f) => sum + (f.isDirectory() ? folderSize(join(dir, f.name)) : statSync(join(dir, f.name)).size), 0)
const mb = (bytes) => `${(bytes / 1024 / 1024).toFixed(2)} MB`
console.log([
  `tabler ${versionOf('@tabler/icons')}: ${Object.keys(tabler).length} icons, ${mb(statSync(join(out, 'tabler.json')).size)}`,
  `flags ${versionOf('flag-icons')}: ${flags.length} flags in 4x3 and 1x1, ${mb(folderSize(join(out, 'flags')))}`,
  `maps ${versionOf('world-atlas')}: ${maps.length} countries and ${regionMaps.length} regions, ${mb(folderSize(mapsOut))}`,
  `left out: ${trimmed.join(', ')}`,
  `public/assets/icons: ${mb(folderSize(out))}`,
].join('\n'))
