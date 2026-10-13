import { existsSync, readFileSync, readdirSync, statSync } from 'node:fs'
import { join } from 'node:path'
import { DOMParser } from '@xmldom/xmldom'
import type { Element as XmlElement } from '@xmldom/xmldom'
import { afterAll, beforeAll, describe, expect, it, vi } from 'vitest'
import type { IconIndex } from '../../src/brand/icons.ts'
import type { FlagJson, IconEntry, LibraryColour, MapEntry, MapJson, TablerJson } from '../../src/brand/iconLibraries.ts'
import {
  LIBRARY_HEX, coloursFor, flagLibrary, flagSvg, iconSvgFor, lucideLibrary, mapLibrary, mapSvg, strokeSvg, tablerLibrary,
} from '../../src/brand/iconLibraries.ts'
import { ALLOWED_HEX } from '../../src/brand/tokens.ts'

const PUBLIC = join(import.meta.dirname, '../../public')
const ICONS = join(PUBLIC, 'assets/icons')
const readJson = <T>(path: string): T => JSON.parse(readFileSync(path, 'utf8')) as T

const tablerJson = readJson<TablerJson>(join(ICONS, 'tabler.json'))
const flagJson = readJson<FlagJson>(join(ICONS, 'flags/flags.json'))
const mapJson = readJson<MapJson>(join(ICONS, 'maps/maps.json'))
const lucide = lucideLibrary(readJson<IconIndex>(join(PUBLIC, 'assets/lucide/index.json')))
const tabler = tablerLibrary(tablerJson)
const maps = mapLibrary(mapJson)

function folderSize(dir: string): number {
  return readdirSync(dir, { withFileTypes: true })
    .reduce((sum, f) => sum + (f.isDirectory() ? folderSize(join(dir, f.name)) : statSync(join(dir, f.name)).size), 0)
}

function parseSvg(svg: string): { root: XmlElement; errors: string[] } {
  const errors: string[] = []
  const doc = new DOMParser({ onError: (level, message) => void errors.push(`${level}: ${message}`) }).parseFromString(svg, 'image/svg+xml')
  const root = doc.documentElement
  if (!root) throw new Error('The SVG has no root element.')
  return { root, errors }
}

function elementsOf(root: XmlElement): XmlElement[] {
  const out: XmlElement[] = [root]
  for (let i = 0; i < out.length; i += 1) {
    const node = out[i]
    if (!node) continue
    for (let c = 0; c < node.childNodes.length; c += 1) {
      const child = node.childNodes[c]
      if (child && child.nodeType === 1) out.push(child as XmlElement)
    }
  }
  return out
}

/** Every paint value in the SVG (fill, stroke and the like), as written. */
function paints(root: XmlElement): Set<string> {
  const found = new Set<string>()
  for (const el of elementsOf(root)) {
    for (const name of ['fill', 'stroke', 'color', 'stop-color', 'flood-color']) {
      const value = el.getAttribute(name)
      if (value !== null) found.add(value)
    }
    expect(el.getAttribute('style'), 'no inline styles').toBeNull()
  }
  return found
}

/** Well-formed SVG whose only paints are "none" and the brand colour asked for. */
function expectBrandSvg(svg: string, colour: LibraryColour, label: string): void {
  const { root, errors } = parseSvg(svg)
  expect(errors, label).toEqual([])
  expect(root.tagName, label).toBe('svg')
  expect(svg, label).not.toMatch(/currentColor/i)
  for (const paint of paints(root)) {
    if (paint === 'none') continue
    expect(paint, label).toBe(`#${LIBRARY_HEX[colour]}`)
    expect(ALLOWED_HEX.has(paint.slice(1).toUpperCase()), label).toBe(true)
  }
}

describe('the generated icon libraries', () => {
  it('has a Tabler index in the Lucide index shape with at least 4,500 outline icons', () => {
    expect(tablerJson.licence).toBe('MIT')
    expect(tablerJson.version).toMatch(/^\d+\.\d+\.\d+$/)
    const names = Object.keys(tablerJson.icons)
    expect(names.length).toBeGreaterThanOrEqual(4500)
    for (const name of names) {
      const icon = tablerJson.icons[name]
      expect(icon?.n.length, name).toBeGreaterThan(0)
      expect(Array.isArray(icon?.t), name).toBe(true)
      expect(typeof icon?.c, name).toBe('string')
    }
    expect(tablerJson.icons.droplet?.t).toContain('water')
  })

  it('has at least 250 flags, each with a name, a region and both SVG files', () => {
    expect(flagJson.flags.length).toBeGreaterThanOrEqual(250)
    const codes = flagJson.flags.map((f) => f.code)
    expect(new Set(codes).size).toBe(codes.length)
    for (const flag of flagJson.flags) {
      expect(flag.name.trim(), flag.code).not.toBe('')
      expect(flag.name, flag.code).not.toBe(flag.code.toUpperCase())
      expect(flag.region.trim(), flag.code).not.toBe('')
      for (const dir of ['4x3', '1x1']) expect(existsSync(join(ICONS, 'flags', dir, `${flag.code}.svg`)), `${dir}/${flag.code}`).toBe(true)
    }
    const name = (code: string): string | undefined => flagJson.flags.find((f) => f.code === code)?.name
    expect(name('eu')).toBe('European Union')
    expect(name('un')).toBe('United Nations')
    expect(name('gb')).toBe('United Kingdom')
    expect([name('gb-eng'), name('gb-sct'), name('gb-wls')]).toEqual(['England', 'Scotland', 'Wales'])
    expect(name('cd')).toBe('Democratic Republic of the Congo')
    expect(codes).not.toContain('gb-nir')
    expect(codes).not.toContain('es-ct')
    expect(codes).not.toContain('xx')
  })

  it('has at least 170 maps, regions first, each file present with its size', () => {
    expect(mapJson.maps.length).toBeGreaterThanOrEqual(170)
    expect(mapJson.maps.slice(0, 9).map((m) => m.code)).toEqual([
      'world', 'africa', 'asia', 'europe', 'latin-america', 'mena', 'sub-saharan-africa', 'south-asia', 'east-asia-pacific',
    ])
    for (const map of mapJson.maps) {
      expect(existsSync(join(PUBLIC, map.file)), map.file).toBe(true)
      expect(map.name.trim(), map.code).not.toBe('')
      expect(Math.max(map.w, map.h), map.code).toBe(1000)
      expect(Math.min(map.w, map.h), map.code).toBeGreaterThan(50)
      const svg = readFileSync(join(PUBLIC, map.file), 'utf8')
      expect(svg, map.code).toContain(`viewBox="0 0 ${map.w} ${map.h}"`)
      expect(svg, map.code).toContain('fill="currentColor"')
      expect(svg, map.code).not.toMatch(/stroke=/)
    }
  })

  it('keeps public/assets/icons well under 15 MB', () => {
    expect(folderSize(ICONS)).toBeLessThan(15 * 1024 * 1024)
    expect(statSync(join(ICONS, 'tabler.json')).size).toBeLessThan(3 * 1024 * 1024)
  })
})

describe('SVG for each library', () => {
  const sample = <T>(list: readonly T[], every: number): T[] => list.filter((_, i) => i % every === 0)

  it('draws every Lucide icon as well-formed SVG in the brand colour asked for, filled dots included', () => {
    for (const entry of lucide.entries) {
      if (entry.kind !== 'stroke') continue
      expectBrandSvg(strokeSvg(entry.nodes, 'green', 24), 'green', entry.id)
    }
    const scatter = lucide.entries.find((e) => e.id === 'chart-scatter')
    if (scatter?.kind !== 'stroke') throw new Error('chart-scatter is missing.')
    for (const colour of [...coloursFor('stroke', false), ...coloursFor('stroke', true)]) expectBrandSvg(strokeSvg(scatter.nodes, colour), colour, colour)
  })

  it('draws Tabler icons as well-formed SVG in brand colours only', () => {
    const filled = tabler.entries.filter((e) => e.kind === 'stroke' && e.nodes.some(([, a]) => Object.values(a).includes('currentColor')))
    expect(filled.length).toBeGreaterThan(10)
    for (const entry of [...sample(tabler.entries, 7), ...filled]) {
      if (entry.kind !== 'stroke') continue
      for (const colour of [...coloursFor('stroke', false), ...coloursFor('stroke', true)]) expectBrandSvg(strokeSvg(entry.nodes, colour, 48), colour, `${entry.id} ${colour}`)
    }
  })

  it('fills every map in the brand colours offered for maps, at the size asked for', () => {
    const colours = [...coloursFor('map', false), ...coloursFor('map', true)]
    expect(colours).toContain('lightGrey')
    for (const entry of maps.entries) {
      if (entry.kind !== 'map') continue
      const text = readFileSync(join(PUBLIC, entry.file), 'utf8')
      for (const colour of colours) expectBrandSvg(mapSvg(text, colour, 300, entry), colour, `${entry.id} ${colour}`)
      const { root } = parseSvg(mapSvg(text, 'navy', 300, entry))
      expect(Math.max(Number(root.getAttribute('width')), Number(root.getAttribute('height'))), entry.id).toBe(300)
    }
  })

  it('keeps flags in their own colours, 4:3, square or round', () => {
    const kenya = readFileSync(join(ICONS, 'flags/4x3/ke.svg'), 'utf8')
    const rect = flagSvg(kenya, 'rect', 400)
    expect(parseSvg(rect).errors).toEqual([])
    expect(rect).toMatch(/<svg[^>]* width="400" height="300">/)
    expect(paints(parseSvg(rect).root)).toEqual(paints(parseSvg(kenya).root))
    const square = readFileSync(join(ICONS, 'flags/1x1/ke.svg'), 'utf8')
    const round = flagSvg(square, 'round', 200)
    expect(parseSvg(round).errors).toEqual([])
    expect(round).toMatch(/<svg[^>]* width="200" height="200">/)
    expect(round).toContain('clip-path="url(#bwb-round)"')
    expect(flagSvg(square, 'square', 200)).not.toContain('clipPath')
  })
})

describe('iconSvgFor', () => {
  beforeAll(() => {
    vi.stubGlobal('fetch', async (url: string) => {
      const path = join(PUBLIC, url.replace(/^\.\//, ''))
      return existsSync(path) ? new Response(readFileSync(path)) : new Response('', { status: 404 })
    })
  })
  afterAll(() => {
    vi.unstubAllGlobals()
  })

  const entry = (list: readonly IconEntry[], id: string): IconEntry => {
    const found = list.find((e) => e.id === id)
    if (!found) throw new Error(`${id} is missing.`)
    return found
  }

  it('draws icons, flags and maps from one call', async () => {
    expectBrandSvg(await iconSvgFor(entry(lucide.entries, 'droplet'), 'navy', 64), 'navy', 'lucide droplet')
    expectBrandSvg(await iconSvgFor(entry(tabler.entries, 'droplet'), 'lime', 64), 'lime', 'tabler droplet')
    const kenya = entry(maps.entries, 'ke') as MapEntry
    const map = await iconSvgFor(kenya, 'lightGrey', 500)
    expectBrandSvg(map, 'lightGrey', 'map ke')
    expect(map).toContain(`height="500"`)
    const flag = await iconSvgFor(entry(flagLibrary(flagJson).entries, 'ke'), 'green', 80)
    expect(flag).toContain('id="flag-icons-ke"')
    expect(flag).not.toContain(`#${LIBRARY_HEX.green}`)
  })

  it('says what failed when a file is missing', async () => {
    const ghost = { ...entry(maps.entries, 'ke'), file: 'assets/icons/maps/none.svg' } as MapEntry
    await expect(iconSvgFor(ghost, 'navy', 100)).rejects.toThrow('The map could not be loaded (404).')
  })
})
