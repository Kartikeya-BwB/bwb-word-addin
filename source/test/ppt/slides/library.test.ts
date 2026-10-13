import { existsSync, readFileSync } from 'node:fs'
import { join } from 'node:path'
import { afterEach, describe, expect, it, vi } from 'vitest'
import pptRibbon from '../../../src/commands/ribbon-powerpoint.json'
import { PPT_SLIDE_HANDLERS } from '../../../src/ppt/slides/handlers.ts'
import { chooseLibrarySlide } from '../../../src/ppt/slides/listSlide.ts'
import type { SlideLibraryIndex } from '../../../src/ppt/slides/library.ts'
import {
  FORMATTING_OPTIONS, SlideLibraryError, bytesToBase64, categoryLabel, clearSlideLibraryCache, findLibrarySlide, libraryBase64,
  loadSlideIndex, parseSlideIndex, searchSlides, slideCategories,
} from '../../../src/ppt/slides/library.ts'

const entry = (id: string, category: string, over: Record<string, unknown> = {}) => ({
  id, title: id.replace(/-/g, ' '), category, description: '', tags: [], slideId: `${256 + id.length}#${id.length * 1000}`,
  thumb: `assets/slides/thumbs/${id}.png`, ...over,
})

const RAW = {
  version: 1,
  file: 'templates/bwb-slides.pptx',
  slides: [
    entry('title', 'title', { title: 'Title slide', tags: ['cover', 'opening'] }),
    entry('agenda', 'agenda', { title: 'Agenda', tags: ['contents', 'sections'] }),
    entry('divider', 'divider', { title: 'Section divider', tags: ['section', 'navy'] }),
    entry('content', 'content', { title: 'Title and bullets', tags: ['bullets', 'text'], description: 'An action title over bullet points.' }),
    entry('key-figures', 'figures', { title: 'Key figures', tags: ['numbers', 'statistics', 'KPI'] }),
    entry('process-three', 'process', { title: 'Process in three steps', tags: ['chevrons', 'steps'] }),
    entry('matrix', 'frameworks', { title: 'Two-by-two matrix', tags: ['quadrants'] }),
  ],
}

afterEach(() => {
  vi.unstubAllGlobals()
  clearSlideLibraryCache()
})

describe('the slide library index', () => {
  it('reads a valid index, lower-casing tags', () => {
    const index = parseSlideIndex(RAW)
    expect(index.slides).toHaveLength(7)
    expect(index.file).toBe('templates/bwb-slides.pptx')
    expect(index.slides[4]?.tags).toEqual(['numbers', 'statistics', 'kpi'])
  })

  it('leaves out broken and repeated entries, and refuses an index that is not one', () => {
    const index = parseSlideIndex({ ...RAW, slides: [...RAW.slides, entry('title', 'title'), { id: 'x' }, null, entry('no-thumb', 'x', { thumb: 7 }), entry('extra', 'x', { tags: 'no' })] })
    expect(index.slides.map((s) => s.id)).toEqual([...RAW.slides.map((s) => s.id), 'extra'])
    expect(index.slides.at(-1)?.tags).toEqual([])
    for (const bad of [null, 'index', { ...RAW, version: 2 }, { ...RAW, file: '' }, { ...RAW, slides: 'none' }, { ...RAW, slides: [] }]) {
      expect(() => parseSlideIndex(bad)).toThrow(SlideLibraryError)
    }
  })

  it('lists categories in catalogue order with readable labels', () => {
    expect(slideCategories(parseSlideIndex(RAW)).map((c) => [c.id, c.label, c.count])).toEqual([
      ['title', 'Title', 1], ['agenda', 'Agenda', 1], ['divider', 'Divider', 1], ['content', 'Content', 1],
      ['figures', 'Figures', 1], ['process', 'Process', 1], ['frameworks', 'Frameworks', 1],
    ])
    expect(categoryLabel('key-figures')).toBe('Key figures')
    expect(categoryLabel('TEAM_PROFILES')).toBe('Team profiles')
    expect(categoryLabel('')).toBe('Other')
  })

  it('searches tags, titles, categories and descriptions, best first, and filters by category', () => {
    const index = parseSlideIndex(RAW)
    expect(searchSlides(index, '').map((s) => s.id)).toEqual(RAW.slides.map((s) => s.id))
    expect(searchSlides(index, 'steps')[0]?.id).toBe('process-three')
    expect(searchSlides(index, 'kpi').map((s) => s.id)).toEqual(['key-figures'])
    expect(searchSlides(index, 'bullet points').map((s) => s.id)).toEqual(['content'])
    expect(searchSlides(index, 'section')[0]?.id).toBe('divider')
    expect(searchSlides(index, '', 'process').map((s) => s.id)).toEqual(['process-three'])
    expect(searchSlides(index, 'agenda', 'process')).toEqual([])
    expect(searchSlides(index, 'zzz')).toEqual([])
  })

  it('finds a library slide by id, id prefix or category', () => {
    const index = parseSlideIndex(RAW)
    expect(findLibrarySlide(index, 'agenda')?.id).toBe('agenda')
    expect(findLibrarySlide(index, 'process')?.id).toBe('process-three')
    expect(findLibrarySlide(index, 'FRAMEWORKS')?.id).toBe('matrix')
    expect(findLibrarySlide(index, 'team')).toBeUndefined()
  })

  it('offers keeping the BwB look first', () => {
    expect(FORMATTING_OPTIONS.map((o) => [o.value, o.label])).toEqual([
      ['KeepSourceFormatting', 'Keep the BwB look'], ['UseDestinationTheme', 'Match this deck'],
    ])
  })
})

describe('loading the library', () => {
  it('fetches the index once, and tries again after a failure', async () => {
    const fetchMock = vi.fn()
      .mockResolvedValueOnce(new Response('missing', { status: 404 }))
      .mockResolvedValue(new Response(JSON.stringify(RAW), { status: 200 }))
    vi.stubGlobal('fetch', fetchMock)
    await expect(loadSlideIndex()).rejects.toThrow('The slide library could not be loaded (404).')
    const index = await loadSlideIndex()
    expect(index.slides).toHaveLength(7)
    await loadSlideIndex()
    expect(fetchMock).toHaveBeenCalledTimes(2)
    expect(fetchMock).toHaveBeenLastCalledWith('./assets/slides/index.json')
  })

  it('fetches the .pptx once and keeps it as base64', async () => {
    const bytes = new Uint8Array(70_000).map((_, i) => (i * 7) % 256)
    const fetchMock = vi.fn().mockImplementation(async () => new Response(bytes, { status: 200 }))
    vi.stubGlobal('fetch', fetchMock)
    const first = await libraryBase64('templates/bwb-slides.pptx')
    const second = await libraryBase64('templates/bwb-slides.pptx')
    expect(first).toBe(Buffer.from(bytes).toString('base64'))
    expect(second).toBe(first)
    expect(fetchMock).toHaveBeenCalledTimes(1)
    expect(fetchMock).toHaveBeenCalledWith('./templates/bwb-slides.pptx')
  })

  it('encodes bytes as base64 the way Node does', () => {
    for (const n of [0, 1, 2, 3, 0x8000, 0x8001]) {
      const bytes = new Uint8Array(n).map((_, i) => (i * 31 + 7) % 256)
      expect(bytesToBase64(bytes)).toBe(Buffer.from(bytes).toString('base64'))
    }
  })
})

const PUBLIC = join(import.meta.dirname, '../../../public')
const INDEX = join(PUBLIC, 'assets/slides/index.json')

describe.skipIf(!existsSync(INDEX))('the published slide library', () => {
  it('lists slides whose thumbnails and file exist, with the slides the slide tools use', () => {
    const index: SlideLibraryIndex = parseSlideIndex(JSON.parse(readFileSync(INDEX, 'utf8')))
    expect(existsSync(join(PUBLIC, index.file)), index.file).toBe(true)
    for (const s of index.slides) {
      expect(existsSync(join(PUBLIC, s.thumb)), s.thumb).toBe(true)
      expect(s.slideId, s.id).toMatch(/^\d*#\d*$/)
    }
    expect(findLibrarySlide(index, 'agenda')).toBeDefined()
    expect(chooseLibrarySlide(index, ['content', 'bullets'])).toBeDefined()
  })
})

interface Entry { fn?: string; items?: Entry[] }

describe('the ribbon handlers for slides', () => {
  it('are exactly the Agenda menu’s functions', () => {
    const groups = (pptRibbon as { groups: Array<{ controls: Entry[] }> }).groups
    const agenda = groups.flatMap((g) => g.controls).find((c) => c.items?.some((i) => i.fn === 'pptAgendaInsert'))
    const fns = (agenda?.items ?? []).flatMap((i) => (i.fn ? [i.fn] : []))
    expect(Object.keys(PPT_SLIDE_HANDLERS).sort()).toEqual([...fns].sort())
    expect(Object.keys(PPT_SLIDE_HANDLERS).sort()).toEqual(['pptAgendaInsert', 'pptSectionMark', 'pptTrackersAdd', 'pptTrackersRemove', 'pptTrackersUpdate'])
  })

  it('say plainly when they run outside PowerPoint', async () => {
    for (const handler of Object.values(PPT_SLIDE_HANDLERS)) await expect(handler()).rejects.toThrow('Open this in PowerPoint to use it.')
  })
})
