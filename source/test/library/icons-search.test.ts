import { readFileSync } from 'node:fs'
import { join } from 'node:path'
import { describe, expect, it } from 'vitest'
import type { IconIndex } from '../../src/brand/icons.ts'
import { ICON_SIZES_CM, resolveIcon, searchIcons } from '../../src/brand/icons.ts'
import type { FlagJson, IconEntry, Library, MapJson, TablerJson } from '../../src/brand/iconLibraries.ts'
import {
  LIBRARIES, SECTORS, altTextFor, boxFor, coloursFor, copyTextFor, flagLibrary, fold, lucideLibrary, mapLibrary, searchLibraries,
  sectorLibrary, sizesFor, tablerLibrary,
} from '../../src/brand/iconLibraries.ts'

const PUBLIC = join(import.meta.dirname, '../../public')
const readJson = <T>(path: string): T => JSON.parse(readFileSync(path, 'utf8')) as T

const index = readJson<IconIndex>(join(PUBLIC, 'assets/lucide/index.json'))
const lucide = lucideLibrary(index)
const sectors = sectorLibrary(index)
const tabler = tablerLibrary(readJson<TablerJson>(join(PUBLIC, 'assets/icons/tabler.json')))
const flags = flagLibrary(readJson<FlagJson>(join(PUBLIC, 'assets/icons/flags/flags.json')))
const maps = mapLibrary(readJson<MapJson>(join(PUBLIC, 'assets/icons/maps/maps.json')))
const all: Library[] = [lucide, sectors, tabler, flags, maps]

const keys = (entries: IconEntry[]): string[] => entries.map((e) => e.key)
const first = (library: Library, query: string): string | undefined => searchLibraries([library], query)[0]?.id

describe('searchLibraries', () => {
  it('finds water in Lucide and Tabler, and solar too', () => {
    for (const library of [lucide, tabler]) {
      const water = searchLibraries([library], 'water')
      expect(water.length, library.info.id).toBeGreaterThan(5)
      const solar = searchLibraries([library], 'solar')
      expect(solar.length, library.info.id).toBeGreaterThan(1)
      expect(solar.slice(0, 3).every((e) => e.id.includes('solar') || e.search.tags.some((t) => t.includes('solar'))), library.info.id).toBe(true)
    }
    expect(first(tabler, 'droplet')).toBe('droplet')
  })

  it('ranks Lucide exactly as searchIcons does', () => {
    for (const query of ['water', 'climate finance', 'solar panel', 'bank', 'drought', 'Gender', 'chart']) {
      const expected = searchIcons(index, query, 40).map((h) => h.name)
      expect(searchLibraries([lucide], query, 40).map((e) => e.id), query).toEqual(expected)
    }
  })

  it('finds Kenya in flags and maps, first in an "All" search', () => {
    expect(first(flags, 'Kenya')).toBe('ke')
    expect(first(maps, 'kenya')).toBe('ke')
    expect(keys(searchLibraries(all, 'Kenya').slice(0, 2)).sort()).toEqual(['flags:ke', 'maps:ke'])
  })

  it('knows places by their codes, short names and other names', () => {
    expect(first(flags, 'UK')).toBe('gb')
    expect(first(flags, 'ke')).toBe('ke')
    expect(first(flags, 'DRC')).toBe('cd')
    expect(first(flags, 'Ivory Coast')).toBe('ci')
    expect(first(flags, "cote d'ivoire")).toBe('ci')
    expect(first(flags, 'Cote d’Ivoire')).toBe('ci')
    expect(first(flags, 'Turkey')).toBe('tr')
    expect(first(flags, 'Wales')).toBe('gb-wls')
    expect(first(flags, 'European Union')).toBe('eu')
    expect(first(flags, 'niger')).toBe('ne')
    expect(first(maps, 'USA')).toBe('us')
    expect(first(maps, 'LAC')).toBe('latin-america')
    expect(first(maps, 'africa')).toBe('africa')
    expect(first(maps, 'south africa')).toBe('za')
  })

  it('finds places in a region after the places named', () => {
    const africa = searchLibraries([flags], 'africa')
    expect(africa[0]?.id).toBe('za')
    expect(africa.length).toBeGreaterThan(40)
  })

  it('finds the BwB sectors by name and by what they cover', () => {
    expect(first(sectors, 'blue economy')).toBe('blue-economy')
    expect(first(sectors, 'ocean')).toBe('blue-economy')
    expect(first(sectors, 'agriculture')).toBe('food')
    expect(first(sectors, 'landmark')).toBe('sovereign')
  })

  it('returns nothing for an empty query or a word nothing has', () => {
    expect(searchLibraries(all, '   ')).toEqual([])
    expect(searchLibraries(all, 'zzqxj')).toEqual([])
  })

  it('folds accents, apostrophes and case', () => {
    expect(fold('Côte d’Ivoire')).toBe('cote-divoire')
    expect(fold('São Tomé & Príncipe')).toBe('sao-tome-and-principe')
    expect(fold('  Solar_Panel ')).toBe('solar-panel')
  })
})

describe('the BwB sector set', () => {
  it('has the 13 website sectors, each a current Lucide icon', () => {
    expect(SECTORS).toHaveLength(13)
    expect(sectors.entries).toHaveLength(13)
    for (const sector of SECTORS) expect(index.icons[sector.icon], sector.icon).toBeDefined()
  })

  it('reaches the same icons through the names the website used before Lucide renamed them', () => {
    expect(resolveIcon(index, 'building-2')).toBe('building-complex')
    expect(resolveIcon(index, 'waves')).toBe('waves-horizontal')
  })
})

describe('colours, sizes and labels', () => {
  const lucideEntry = lucide.entries.find((e) => e.id === 'droplet')
  const flagEntry = flags.entries.find((e) => e.id === 'ke')
  const mapEntry = maps.entries.find((e) => e.id === 'cl')
  if (!lucideEntry || !flagEntry || !mapEntry) throw new Error('Sample entries are missing.')

  it('offers green, navy and charcoal on white, white and lime on navy, and nothing for flags', () => {
    expect(coloursFor('stroke', false)).toEqual(['green', 'navy', 'charcoal'])
    expect(coloursFor('stroke', true)).toEqual(['white', 'lime'])
    expect(coloursFor('map', false)).toEqual(['green', 'navy', 'charcoal', 'lightGrey'])
    expect(coloursFor('map', true)).toEqual(['white', 'lime'])
    expect(coloursFor('flag', false)).toEqual([])
    expect(coloursFor('flag', true)).toEqual([])
  })

  it('sizes icons with the BwB sizes in Word and in points in PowerPoint', () => {
    const word = sizesFor('stroke', 'word')
    expect(word.map((s) => s.id)).toEqual(['inline', 'callout', 'marker', 'panel'])
    expect(word.map((s) => Math.round(s.pt * 100) / 100)).toEqual(Object.values(ICON_SIZES_CM).map((cm) => Math.round(((cm * 72) / 2.54) * 100) / 100))
    expect(word[1]?.label).toBe('Callout, 0.6 cm')
    expect(sizesFor('map', 'word').map((s) => s.id)).toEqual(['small', 'medium', 'large'])
    expect(sizesFor('flag', 'powerpoint').map((s) => s.pt)).toEqual([24, 48, 96])
  })

  it('keeps flags at 4:3 and maps at their own aspect', () => {
    expect(boxFor(lucideEntry, 30)).toEqual({ widthPt: 30, heightPt: 30 })
    expect(boxFor(flagEntry, 40)).toEqual({ widthPt: 40, heightPt: 30 })
    expect(boxFor(flagEntry, 40, 'round')).toEqual({ widthPt: 40, heightPt: 40 })
    const chile = boxFor(mapEntry, 200)
    expect(chile.heightPt).toBe(200)
    expect(chile.widthPt).toBeLessThan(60)
  })

  it('writes alt text and copy text in the BwB pattern', () => {
    expect(altTextFor(lucideEntry)).toBe('Icon: droplet')
    expect(altTextFor(flagEntry)).toBe('Flag: Kenya')
    expect(altTextFor(mapEntry)).toBe('Map: Chile')
    const blue = sectors.entries.find((e) => e.id === 'blue-economy')
    if (!blue) throw new Error('The blue economy sector is missing.')
    expect(altTextFor(blue)).toBe('Icon: Blue economy')
    expect(copyTextFor(blue)).toBe('waves-horizontal')
    expect(copyTextFor(flagEntry)).toBe('Kenya')
  })

  it('lists the libraries with Lucide first and a licence line for each', () => {
    expect(LIBRARIES.map((l) => l.id)).toEqual(['lucide', 'sectors', 'tabler', 'flags', 'maps'])
    for (const library of LIBRARIES) expect(library.licence, library.id).toMatch(/(ISC|MIT) licence/)
  })
})
