// Builds the geometry for the chart builder's data maps:
//   public/assets/charts/world.json   every country as one SVG path on the Equal Earth projection, keyed by ISO 3166
//                                      alpha-2 and alpha-3 codes, with names, aliases, regions and region frames
//   public/assets/charts/LICENSE.txt  world-atlas's ISC licence and the Natural Earth credit
// Boundaries follow the icon library (scripts/make-icon-libraries.mjs): Crimea with Ukraine, Western Sahara split
// along 27 degrees 40 minutes north, Somaliland with Somalia and Northern Cyprus with Cyprus.
// The output is committed; run `node scripts/make-chart-maps.mjs` again after updating world-atlas.
import { mkdirSync, readFileSync, statSync, writeFileSync } from 'node:fs'
import { createRequire } from 'node:module'
import { dirname, join } from 'node:path'
import { fileURLToPath } from 'node:url'
import { geoContains, geoEqualEarth, geoIdentity, geoPath, geoStream } from 'd3-geo'

const require = createRequire(import.meta.url)
const { feature, merge } = require('topojson-client')

const root = join(dirname(fileURLToPath(import.meta.url)), '..')
const modules = join(root, 'node_modules')
const out = join(root, 'public', 'assets', 'charts')
const readJson = (path) => JSON.parse(readFileSync(path, 'utf8'))
const versionOf = (pkg) => readJson(join(modules, pkg, 'package.json')).version

/** Width of the projected plane; coordinates are whole units, about 10 km at the equator. */
const WIDTH = 8000
/** Simplification of the shared borders, in degrees, before projecting. */
const TOLERANCE = 0.012
/** Simplification again after projecting, in plane units. */
const PLANE_TOLERANCE = 0.3
/** Smallest island kept beside a country's main shape, in square units. */
const MIN_RING_AREA = 24

// ---- ISO 3166-1 alpha-2 to alpha-3, with Kosovo's widely used XKX ----
const ALPHA3 = Object.fromEntries(`
AD:AND AE:ARE AF:AFG AG:ATG AI:AIA AL:ALB AM:ARM AO:AGO AQ:ATA AR:ARG AS:ASM AT:AUT AU:AUS AW:ABW AX:ALA AZ:AZE BA:BIH
BB:BRB BD:BGD BE:BEL BF:BFA BG:BGR BH:BHR BI:BDI BJ:BEN BL:BLM BM:BMU BN:BRN BO:BOL BQ:BES BR:BRA BS:BHS BT:BTN BV:BVT
BW:BWA BY:BLR BZ:BLZ CA:CAN CC:CCK CD:COD CF:CAF CG:COG CH:CHE CI:CIV CK:COK CL:CHL CM:CMR CN:CHN CO:COL CR:CRI CU:CUB
CV:CPV CW:CUW CX:CXR CY:CYP CZ:CZE DE:DEU DJ:DJI DK:DNK DM:DMA DO:DOM DZ:DZA EC:ECU EE:EST EG:EGY EH:ESH ER:ERI ES:ESP
ET:ETH FI:FIN FJ:FJI FK:FLK FM:FSM FO:FRO FR:FRA GA:GAB GB:GBR GD:GRD GE:GEO GF:GUF GG:GGY GH:GHA GI:GIB GL:GRL GM:GMB
GN:GIN GP:GLP GQ:GNQ GR:GRC GS:SGS GT:GTM GU:GUM GW:GNB GY:GUY HK:HKG HM:HMD HN:HND HR:HRV HT:HTI HU:HUN ID:IDN IE:IRL
IL:ISR IM:IMN IN:IND IO:IOT IQ:IRQ IR:IRN IS:ISL IT:ITA JE:JEY JM:JAM JO:JOR JP:JPN KE:KEN KG:KGZ KH:KHM KI:KIR KM:COM
KN:KNA KP:PRK KR:KOR KW:KWT KY:CYM KZ:KAZ LA:LAO LB:LBN LC:LCA LI:LIE LK:LKA LR:LBR LS:LSO LT:LTU LU:LUX LV:LVA LY:LBY
MA:MAR MC:MCO MD:MDA ME:MNE MF:MAF MG:MDG MH:MHL MK:MKD ML:MLI MM:MMR MN:MNG MO:MAC MP:MNP MQ:MTQ MR:MRT MS:MSR MT:MLT
MU:MUS MV:MDV MW:MWI MX:MEX MY:MYS MZ:MOZ NA:NAM NC:NCL NE:NER NF:NFK NG:NGA NI:NIC NL:NLD NO:NOR NP:NPL NR:NRU NU:NIU
NZ:NZL OM:OMN PA:PAN PE:PER PF:PYF PG:PNG PH:PHL PK:PAK PL:POL PM:SPM PN:PCN PR:PRI PS:PSE PT:PRT PW:PLW PY:PRY QA:QAT
RE:REU RO:ROU RS:SRB RU:RUS RW:RWA SA:SAU SB:SLB SC:SYC SD:SDN SE:SWE SG:SGP SH:SHN SI:SVN SJ:SJM SK:SVK SL:SLE SM:SMR
SN:SEN SO:SOM SR:SUR SS:SSD ST:STP SV:SLV SX:SXM SY:SYR SZ:SWZ TC:TCA TD:TCD TF:ATF TG:TGO TH:THA TJ:TJK TK:TKL TL:TLS
TM:TKM TN:TUN TO:TON TR:TUR TT:TTO TV:TUV TW:TWN TZ:TZA UA:UKR UG:UGA UM:UMI US:USA UY:URY UZ:UZB VA:VAT VC:VCT VE:VEN
VG:VGB VI:VIR VN:VNM VU:VUT WF:WLF WS:WSM XK:XKX YE:YEM YT:MYT ZA:ZAF ZM:ZMB ZW:ZWE`.trim().split(/\s+/).map((pair) => pair.split(':')))

// Other names people paste: short forms, older names, and the styles used in development statistics.
const ALIASES = Object.fromEntries(`
AE: UAE; Emirates
BA: Bosnia; Bosnia-Herzegovina
BN: Brunei Darussalam
BO: Bolivia (Plurinational State of); Plurinational State of Bolivia
BS: The Bahamas; Bahamas, The
CD: DRC; DR Congo; Congo-Kinshasa; Congo (Kinshasa); Congo, Dem. Rep.; Democratic Republic of Congo; Dem. Rep. Congo
CF: CAR
CG: Congo; Congo-Brazzaville; Congo (Brazzaville); Congo, Rep.; Republic of Congo
CI: Ivory Coast; Cote d'Ivoire
CN: PRC; People's Republic of China
CV: Cabo Verde; Cape Verde
CZ: Czech Republic
EG: Egypt, Arab Rep.; Arab Republic of Egypt
FM: Federated States of Micronesia; Micronesia, Fed. Sts.; Micronesia (Federated States of)
GB: UK; Britain; Great Britain; United Kingdom of Great Britain and Northern Ireland
GM: The Gambia; Gambia, The
HK: Hong Kong SAR, China; Hong Kong, China
IR: Iran, Islamic Rep.; Islamic Republic of Iran; Iran (Islamic Republic of)
KG: Kyrgyz Republic
KN: St. Kitts and Nevis
KP: DPRK; Korea, Dem. People's Rep.; Democratic People's Republic of Korea
KR: Korea; Korea, Rep.; Republic of Korea
LA: Lao PDR; Lao People's Democratic Republic
LC: St. Lucia
MD: Republic of Moldova
MF: St. Martin (French part)
MK: Macedonia; Republic of North Macedonia
MM: Burma
MO: Macao SAR, China; Macau
NL: Holland
PG: PNG
PS: Palestine; State of Palestine; West Bank and Gaza; West Bank; Gaza
RU: Russian Federation
SA: KSA
SK: Slovak Republic
ST: Sao Tome and Principe
SX: Sint Maarten (Dutch part)
SY: Syrian Arab Republic
SZ: Swaziland
TL: East Timor
TR: Turkey; Turkiye
TW: Taiwan, China; Chinese Taipei
TZ: United Republic of Tanzania; Tanzania, United Republic of
US: USA; U.S.; America; United States of America
VA: Holy See; Vatican
VC: St. Vincent and the Grenadines
VE: Venezuela, RB; Venezuela (Bolivarian Republic of); Bolivarian Republic of Venezuela
VG: British Virgin Islands
VI: Virgin Islands (U.S.); US Virgin Islands
VN: Viet Nam
YE: Yemen, Rep.; Republic of Yemen`.trim().split('\n').map((line) => {
  const [code, names] = line.split(': ')
  return [code, names.split('; ')]
}))

// World Bank regions, as in the icon library, with territories placed by geography.
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

// ---- Names, as in the icon library ----
const longName = new Intl.DisplayNames('en-GB', { type: 'region' })
const shortName = new Intl.DisplayNames('en-GB', { type: 'region', style: 'short' })
const NAME_FIXES = { CD: 'Democratic Republic of the Congo', CG: 'Republic of the Congo', MM: 'Myanmar', HK: 'Hong Kong', MO: 'Macao' }
const nameOf = (code) => (NAME_FIXES[code] ?? longName.of(code)).replace(/ & /g, ' and ')
const flagCountries = readJson(join(modules, 'flag-icons', 'country.json'))
const continentOf = new Map(flagCountries.map((c) => [c.code.toUpperCase(), c.continent]))
const flagName = new Map(flagCountries.map((c) => [c.code.toUpperCase(), c.name]))

function aliasesOf(code, name) {
  const candidates = [longName.of(code), shortName.of(code), flagName.get(code) ?? '', ...(ALIASES[code] ?? [])]
  return [...new Set(candidates.map((a) => a.replace(/ & /g, ' and ').replace(/ - /g, '-')))]
    .filter((a) => a && a !== name && a.toUpperCase() !== code)
}

const regionByCode = new Map(Object.entries(REGIONS).flatMap(([region, codes]) => codes.split(' ').map((c) => [c, region])))

// ---- Simplifying shared borders, so neighbours keep identical edges ----
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

function simplifyRing(ring, tolerance) {
  if (ring.length < 4 || tolerance <= 0) return ring
  const [x0, y0] = ring[0]
  const distances = ring.map(([x, y]) => Math.hypot(x - x0, y - y0))
  const split = distances.reduce((best, d, i) => (d > distances[best] ? i : best), 0)
  return [...simplify(ring.slice(0, split + 1), tolerance), ...simplify([...ring.slice(split), ring[0]], tolerance).slice(1, -1)]
}

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

/** Cuts a lon/lat geometry to a box, as the icon library does for Western Sahara. */
function cutLonLat(geometry, [[x0, y0], [x1, y1]]) {
  const polygons = []
  let [rings, ring] = [[], []]
  const sink = {
    point: (x, y) => void ring.push([x, -y]), lineStart: () => void (ring = []), sphere: () => {},
    lineEnd: () => void (ring.length > 2 && rings.push([...ring, ring[0]])),
    polygonStart: () => void (rings = []), polygonEnd: () => void (rings.length && polygons.push(rings)),
  }
  geoStream(geometry, geoIdentity().reflectY(true).clipExtent([[x0, -y1], [x1, -y0]]).stream(sink))
  return { type: 'MultiPolygon', coordinates: polygons }
}

// ---- Countries from Natural Earth 1:50m ----
const atlas = join(modules, 'world-atlas')
const topology = readJson(join(atlas, 'countries-50m.json'))
const UNIT_CODES = { Somaliland: 'SO', 'N. Cyprus': 'CY', 'Siachen Glacier': 'IN', Kosovo: 'XK' }
const SKIP_UNITS = new Set(['Ashmore and Cartier Is.', 'Indian Ocean Ter.', 'Antarctica'])
const units = new Map()
for (const g of topology.objects.countries.geometries) {
  const name = g.properties.name
  const code = SKIP_UNITS.has(name) ? null : UNIT_CODES[name] ?? new Intl.Locale(`und-${g.id}`).region
  if (code && code !== 'AQ') units.set(code, [...(units.get(code) ?? []), g])
}
const russia = units.get('RU')?.[0]
const crimea = russia?.arcs.findIndex((polygon) => geoContains(feature(topology, { type: 'Polygon', arcs: polygon }), [34.1, 45.0])) ?? -1
if (russia && crimea >= 0) units.get('UA')?.push({ type: 'Polygon', arcs: russia.arcs.splice(crimea, 1)[0] })

const simple = simplifyTopology(topology, TOLERANCE)
const shapes = new Map([...units].map(([code, geometries]) => [code, merge(simple, geometries)]))
const saharan = merge(simple, [...(units.get('MA') ?? []), ...(units.get('EH') ?? [])])
shapes.set('MA', cutLonLat(saharan, [[-180, 27.6667], [180, 90]]))
shapes.set('EH', cutLonLat(saharan, [[-180, -90], [180, 27.6667]]))

// Island states and territories too small for the 1:50m data come from 1:10m. None has a land border, so no shared
// edge can disagree with a neighbour's.
const ISLANDS = new Set(['BL', 'CK', 'GG', 'MS', 'MV', 'NF', 'NR', 'PN', 'TV', 'UM'])
const detailed = readJson(join(atlas, 'countries-10m.json'))
const islands = new Map()
for (const g of detailed.objects.countries.geometries) {
  const code = g.id ? new Intl.Locale(`und-${g.id}`).region : null
  if (code && ISLANDS.has(code) && !shapes.has(code)) islands.set(code, [...(islands.get(code) ?? []), g])
}
for (const [code, geometries] of islands) shapes.set(code, merge(detailed, geometries))

// ---- Projection: the whole sphere fitted to WIDTH, so every region shares one plane ----
const projection = geoEqualEarth().fitWidth(WIDTH, { type: 'Sphere' }).precision(0.1)
const [, [, sphereBottom]] = geoPath(projection).bounds({ type: 'Sphere' })
const HEIGHT = Math.round(sphereBottom)

const ringArea = (r) => r.reduce((sum, [x, y], i) => sum + x * r[(i + 1) % r.length][1] - r[(i + 1) % r.length][0] * y, 0) / 2

function projectedRings(geometry) {
  const rings = []
  let ring = []
  const context = { moveTo: (x, y) => rings.push((ring = [[x, y]])), lineTo: (x, y) => void ring.push([x, y]), closePath() {}, arc() {} }
  geoPath(projection, context)(geometry)
  return rings
}

const numbers = (values) => values.map((n, i) => (i && n >= 0 ? ` ${n}` : `${n}`)).join('')

/** One path: whole-unit points, lightly simplified again in the plane, specks dropped but never the main shape. */
function pathOf(geometry) {
  const rings = projectedRings(geometry)
    .map((r) => simplifyRing(r, PLANE_TOLERANCE).map(([x, y]) => [Math.round(x), Math.round(y)]))
    .map((r) => r.filter((p, i) => i === 0 || p[0] !== r[i - 1][0] || p[1] !== r[i - 1][1]))
    .map((r) => (r.length > 1 && r[0][0] === r.at(-1)[0] && r[0][1] === r.at(-1)[1] ? r.slice(0, -1) : r))
    .filter((r) => r.length >= 3)
  if (!rings.length) return null
  const areas = rings.map((r) => Math.abs(ringArea(r)))
  const main = areas.indexOf(Math.max(...areas))
  const kept = rings.filter((_, i) => i === main || areas[i] >= MIN_RING_AREA)
  const all = kept.flat()
  const box = [Math.min(...all.map((p) => p[0])), Math.min(...all.map((p) => p[1])), Math.max(...all.map((p) => p[0])), Math.max(...all.map((p) => p[1]))]
  const d = kept.map((r) => `M${r[0][0]} ${r[0][1]}l${numbers(r.slice(1).flatMap(([x, y], i) => [x - r[i][0], y - r[i][1]]))}z`).join('')
  return { d, box }
}

const countries = []
const unknownRegion = []
for (const code of Object.keys(ALPHA3).sort()) {
  const name = nameOf(code)
  const region = regionByCode.get(code)
  if (!region) unknownRegion.push(code)
  const drawn = shapes.has(code) ? pathOf(shapes.get(code)) : null
  countries.push({
    a2: code, a3: ALPHA3[code], n: name, al: aliasesOf(code, name), r: region ?? '', c: continentOf.get(code) ?? '',
    d: drawn?.d ?? '', b: drawn?.box ?? null,
  })
}
const missingCodes = [...shapes.keys()].filter((c) => !ALPHA3[c])
if (missingCodes.length) throw new Error(`No alpha-3 code for ${missingCodes.join(', ')}: add them to ALPHA3.`)

// ---- Region frames: a box in degrees, projected, so each region fills the chart ----
function frameOf([[lon0, lat0], [lon1, lat1]]) {
  const points = []
  for (let lon = lon0; lon <= lon1; lon += 0.5) points.push([lon, lat0], [lon, lat1])
  for (let lat = lat0; lat <= lat1; lat += 0.5) points.push([lon0, lat], [lon1, lat])
  const xy = points.map((p) => projection(p))
  return [Math.min(...xy.map((p) => p[0])), Math.min(...xy.map((p) => p[1])), Math.max(...xy.map((p) => p[0])), Math.max(...xy.map((p) => p[1]))]
    .map((v) => Math.round(v))
}
const membersWhere = (test) => countries.filter((c) => test(c)).map((c) => c.a2)
const FRAMES = [
  { id: 'world', name: 'World', box: [[-180, -57], [180, 84]], members: membersWhere(() => true) },
  { id: 'africa', name: 'Africa', box: [[-26, -36], [58, 38]], members: membersWhere((c) => c.c === 'Africa') },
  { id: 'europe', name: 'Europe', box: [[-25, 34], [45, 71.5]], members: membersWhere((c) => c.c === 'Europe' && c.a2 !== 'RU') },
  { id: 'asia', name: 'Asia', box: [[25, -11], [150, 56]], members: membersWhere((c) => c.c === 'Asia') },
  { id: 'latin-america', name: 'Latin America and the Caribbean', box: [[-118, -56], [-33, 33]], members: membersWhere((c) => c.r === 'Latin America and the Caribbean') },
  { id: 'mena', name: 'Middle East and North Africa', box: [[-18, 11], [64, 42]], members: membersWhere((c) => c.r === 'Middle East and North Africa') },
  { id: 'sub-saharan-africa', name: 'Sub-Saharan Africa', box: [[-26, -36], [58, 28]], members: membersWhere((c) => c.r === 'Sub-Saharan Africa') },
]
const regions = FRAMES.map((f) => ({ id: f.id, name: f.name, box: frameOf(f.box), members: f.members }))

mkdirSync(out, { recursive: true })
const world = {
  version: versionOf('world-atlas'),
  source: 'Natural Earth 4.1.0 via world-atlas',
  licence: 'ISC (world-atlas); Natural Earth data is in the public domain',
  credit: 'Boundaries: Natural Earth',
  projection: 'Equal Earth',
  width: WIDTH,
  height: HEIGHT,
  countries,
  regions,
}
writeFileSync(join(out, 'world.json'), JSON.stringify(world))
writeFileSync(join(out, 'LICENSE.txt'),
  `${readFileSync(join(atlas, 'LICENSE'), 'utf8')}\nThe map data comes from Natural Earth (naturalearthdata.com), which is in the public domain.\n`)

const kb = (path) => `${(statSync(path).size / 1024).toFixed(0)} KB`
console.log([
  `world-atlas ${versionOf('world-atlas')}: ${countries.filter((c) => c.d).length} countries drawn, ${countries.filter((c) => !c.d).length} named only`,
  `named only: ${countries.filter((c) => !c.d).map((c) => c.a2).join(' ')}`,
  unknownRegion.length ? `no World Bank region: ${unknownRegion.join(' ')}` : 'every country has a region',
  `public/assets/charts/world.json: ${kb(join(out, 'world.json'))}`,
].join('\n'))
