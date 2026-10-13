import { existsSync, readFileSync, statSync } from 'node:fs'
import { join } from 'node:path'
import { strFromU8, unzipSync, type Unzipped } from 'fflate'
import { describe, expect, it } from 'vitest'
import { CONTENT_AREA, SLIDE, TITLE_AREA, type Area } from '../../src/brand/slideGrid.ts'
import { COLOURS } from '../../src/brand/tokens.ts'

// The slide library: public/assets/slides/index.json, the template it points at and the starter deck, all written by
// scripts/make-slide-library.py. The add-in inserts a slide with insertSlidesFromBase64 and the slide's
// "<sldId>#<creationId>" id, so every id in the index must exist in the pptx. A new deck starts from the starter deck.

interface SlideEntry {
  id: string
  title: string
  category: string
  description: string
  tags: string[]
  slideId: string
  thumb: string
}

interface SlideIndex {
  version: number
  file: string
  starterDeck: string
  slides: SlideEntry[]
}

const PUBLIC = join(import.meta.dirname, '../../public')
const index = JSON.parse(readFileSync(join(PUBLIC, 'assets/slides/index.json'), 'utf8')) as SlideIndex
const library = unzipSync(readFileSync(join(PUBLIC, index.file)))
const deck = unzipSync(readFileSync(join(PUBLIC, index.starterDeck)))
const EMU_PER_POINT = 12700
const CATEGORIES = ['Structure', 'Messages', 'Data', 'Visuals', 'People', 'Closing']
const LAYOUTS = ['BwB Title', 'BwB Section divider', 'BwB Title and content', 'BwB Two content', 'BwB Title only',
  'BwB Blank', 'BwB Closing']

function part(zip: Unzipped, name: string): string {
  const data = zip[name]
  if (!data) throw new Error(`${name} is missing from the package`)
  return strFromU8(data)
}

function attr(tag: string, name: string): string {
  return new RegExp(`\\s${name}="([^"]*)"`).exec(tag)?.[1] ?? ''
}

/** The targets of a part's relationships by id, as package paths. */
function relationships(zip: Unzipped, relsName: string, base: string): Map<string, string> {
  const targets = new Map<string, string>()
  for (const [rel] of part(zip, relsName).matchAll(/<Relationship [^>]*>/g)) {
    const target = attr(rel, 'Target')
    const parts = target.startsWith('/') ? [] : base.split('/')
    for (const step of target.replace(/^\//, '').split('/')) {
      if (step === '..') parts.pop()
      else parts.push(step)
    }
    targets.set(attr(rel, 'Id'), parts.join('/'))
  }
  return targets
}

/** Maps each <p:sldId id> in ppt/presentation.xml, in order, to its slide part. */
function slideParts(zip: Unzipped): Map<string, string> {
  const targets = relationships(zip, 'ppt/_rels/presentation.xml.rels', 'ppt')
  const slides = new Map<string, string>()
  for (const [tag] of part(zip, 'ppt/presentation.xml').matchAll(/<p:sldId [^>]*>/g)) {
    slides.set(attr(tag, 'id'), targets.get(attr(tag, 'r:id')) ?? '')
  }
  return slides
}

function layoutName(zip: Unzipped, layoutPart: string): string {
  return /<p:cSld name="([^"]+)"/.exec(part(zip, layoutPart))?.[1] ?? layoutPart
}

function layoutNames(zip: Unzipped): string[] {
  return Object.keys(zip)
    .filter((name) => /^ppt\/slideLayouts\/slideLayout\d+\.xml$/.test(name))
    .map((name) => layoutName(zip, name))
    .sort()
}

/** The layout each slide uses, in slide order. */
function slideLayouts(zip: Unzipped): string[] {
  return [...slideParts(zip).values()].map((slidePart) => {
    const rels = relationships(zip, slidePart.replace(/([^/]+)$/, '_rels/$1.rels'), 'ppt/slides')
    const layout = [...rels.values()].find((target) => target.startsWith('ppt/slideLayouts/')) ?? ''
    return layoutName(zip, layout)
  })
}

/** The position and size of the placeholder with these attributes, in EMU. */
function placeholderBox(xml: string, ph: string): number[] {
  const shape = xml.split('<p:sp>').find((s) => s.includes(`<p:ph ${ph}/>`)) ?? ''
  const m = /<a:off x="(\d+)" y="(\d+)"\/><a:ext cx="(\d+)" cy="(\d+)"\/>/.exec(shape)
  return m ? m.slice(1).map(Number) : []
}

function emu(area: Area): number[] {
  return [area.left, area.top, area.width, area.height].map((pt) => Math.round(pt * EMU_PER_POINT))
}

function embeddedFonts(zip: Unzipped): string[] {
  const fonts = Object.keys(zip).filter((name) => name.startsWith('ppt/fonts/') || /\.(fntdata|odttf|ttf|otf)$/i.test(name))
  return part(zip, 'ppt/presentation.xml').includes('embeddedFont') ? [...fonts, 'embeddedFontLst'] : fonts
}

describe('the slide library index', () => {
  it('points at the template and the starter deck, and lists every slide in the template', () => {
    expect(index.version).toBe(1)
    expect(index.file).toBe('templates/bwb-slides.pptx')
    expect(index.starterDeck).toBe('templates/bwb-deck.pptx')
    expect(existsSync(join(PUBLIC, index.file))).toBe(true)
    expect(existsSync(join(PUBLIC, index.starterDeck))).toBe(true)
    expect(index.slides.length).toBeGreaterThanOrEqual(20)
    expect(slideParts(library).size).toBe(index.slides.length)
  })

  it('has unique ids, and a known category, a title, a description and tags for every slide', () => {
    const ids = index.slides.map((s) => s.id)
    expect(new Set(ids).size).toBe(ids.length)
    expect(new Set(index.slides.map((s) => s.slideId)).size).toBe(ids.length)
    for (const s of index.slides) {
      expect(s.id).toMatch(/^[a-z0-9]+(-[a-z0-9]+)*$/)
      expect(CATEGORIES, s.id).toContain(s.category)
      expect(s.title.trim(), s.id).not.toBe('')
      expect(s.description.trim(), s.id).not.toBe('')
      expect(s.tags.length, s.id).toBeGreaterThan(1)
    }
  })

  it('has a 480 px PNG thumbnail for every slide', () => {
    for (const s of index.slides) {
      expect(s.thumb).toBe(`assets/slides/thumbs/${s.id}.png`)
      const path = join(PUBLIC, s.thumb)
      expect(existsSync(path), s.thumb).toBe(true)
      const png = readFileSync(path)
      expect(png.subarray(1, 4).toString('latin1'), s.thumb).toBe('PNG')
      expect(png.readUInt32BE(16), s.thumb).toBe(480)
    }
  })

  it('gives every slide the sldId and creationId it has in the pptx', () => {
    const slides = slideParts(library)
    for (const s of index.slides) {
      const m = /^(\d+)#(\d+)$/.exec(s.slideId)
      expect(m, s.slideId).not.toBeNull()
      const [, sldId = '', creationId = ''] = m ?? []
      const slidePart = slides.get(sldId)
      expect(slidePart, `${s.id}: no <p:sldId id="${sldId}"> in ppt/presentation.xml`).toBeTruthy()
      const created = /<p14:creationId [^>]*val="(\d+)"/.exec(part(library, slidePart ?? ''))?.[1]
      expect(created, `${s.id}: creationId of ${slidePart ?? ''}`).toBe(creationId)
    }
  })
})

describe('the BwB slide template', () => {
  it('is 16:9, with the title and text placeholders on the BwB grid', () => {
    const size = emu(SLIDE)
    expect(part(library, 'ppt/presentation.xml')).toContain(`<p:sldSz cx="${size[2] ?? 0}" cy="${size[3] ?? 0}"/>`)
    const master = part(library, 'ppt/slideMasters/slideMaster1.xml')
    expect(placeholderBox(master, 'type="title"')).toEqual(emu(TITLE_AREA))
    expect(placeholderBox(master, 'type="body" idx="1"')).toEqual(emu(CONTENT_AREA))
  })

  it('has the seven BwB layouts', () => {
    expect(layoutNames(library)).toEqual([...LAYOUTS].sort())
  })

  it('uses the BwB colours and fonts in its theme, and embeds no fonts', () => {
    const theme = part(library, 'ppt/theme/theme1.xml')
    const slots: Array<[string, string]> = [['dk1', COLOURS.black], ['lt1', COLOURS.white], ['dk2', COLOURS.navy],
      ['lt2', COLOURS.offWhite], ['accent1', COLOURS.navy], ['accent2', COLOURS.green], ['accent3', COLOURS.green50],
      ['accent4', COLOURS.charcoal], ['accent5', COLOURS.lime], ['accent6', COLOURS.lightGrey],
      ['hlink', COLOURS.green], ['folHlink', COLOURS.charcoal]]
    for (const [slot, hex] of slots) expect(theme).toContain(`<a:${slot}><a:srgbClr val="${hex}"/></a:${slot}>`)
    expect(theme).toContain('<a:majorFont><a:latin typeface="Candara"/>')
    expect(theme).toContain('<a:minorFont><a:latin typeface="Aptos"/>')
    expect(embeddedFonts(library)).toEqual([])
  })

  it('has a native chart on the chart slide, with the workbook behind it for Edit Data', () => {
    const entry = index.slides.find((s) => s.id === 'chart-and-text')
    const slidePart = slideParts(library).get(entry?.slideId.split('#')[0] ?? '') ?? ''
    const frame = /<p:graphicFrame>(?:(?!<\/p:graphicFrame>).)*name="Chart"(?:(?!<\/p:graphicFrame>).)*<\/p:graphicFrame>/s
      .exec(part(library, slidePart))?.[0] ?? ''
    const chartRid = /<c:chart [^>]*r:id="([^"]+)"/.exec(frame)?.[1] ?? ''
    const chartPart = relationships(library, slidePart.replace(/([^/]+)$/, '_rels/$1.rels'), 'ppt/slides').get(chartRid) ?? ''
    expect(chartPart).toMatch(/^ppt\/charts\/chart\d+\.xml$/)
    const chart = part(library, chartPart)
    expect(chart).toContain('<c:autoTitleDeleted val="1"/>')
    expect(chart).toContain(`<a:srgbClr val="${COLOURS.green}"/>`)
    expect(chart).toContain(`<a:srgbClr val="${COLOURS.green50}"/>`)
    expect(chart).toContain('<a:latin typeface="Aptos"/>')
    const dataRid = /<c:externalData r:id="([^"]+)"/.exec(chart)?.[1] ?? ''
    const workbook = relationships(library, chartPart.replace(/([^/]+)$/, '_rels/$1.rels'), 'ppt/charts').get(dataRid) ?? ''
    expect(workbook).toMatch(/^ppt\/embeddings\/.+\.xlsx$/)
    expect(library[workbook]?.length ?? 0).toBeGreaterThan(0)
  })
})

describe('the starter deck', () => {
  it('has exactly two slides: a title slide and a title and content slide', () => {
    expect(slideParts(deck).size).toBe(2)
    expect(slideLayouts(deck)).toEqual(['BwB Title', 'BwB Title and content'])
  })

  it('leaves the placeholders empty, so PowerPoint shows the layout prompts such as the action title', () => {
    const [titleSlide = '', contentSlide = ''] = [...slideParts(deck).values()].map((name) => part(deck, name))
    for (const xml of [titleSlide, contentSlide]) {
      expect(xml.replace(/<a:fld [^>]*type="slidenum".*?<\/a:fld>/s, '')).not.toMatch(/<a:t>[^<]+<\/a:t>/)
    }
    expect(titleSlide).toContain('<p:ph type="ctrTitle"/>')
    expect(contentSlide).toContain('<p:ph type="title"/>')
    const layout = Object.keys(deck).find((name) => /^ppt\/slideLayouts\/slideLayout\d+\.xml$/.test(name) &&
      layoutName(deck, name) === 'BwB Title and content') ?? ''
    expect(part(deck, layout)).toMatch(/<a:t>Action title: [^<]+<\/a:t>/)
  })

  it('has the same layouts as the slide library', () => {
    expect(layoutNames(deck)).toEqual(layoutNames(library))
  })

  it('embeds no fonts, holds no photos and stays under 150 KB', () => {
    expect(embeddedFonts(deck)).toEqual([])
    expect(Object.keys(deck).filter((name) => /\.(jpe?g)$/i.test(name))).toEqual([])
    expect(statSync(join(PUBLIC, index.starterDeck)).size).toBeLessThan(150 * 1000)
  })
})
