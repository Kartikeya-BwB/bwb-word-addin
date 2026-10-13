// The BwB slide library for tests: the index and .pptx served by a stubbed fetch, and the library's agenda, text and
// divider slides built in the fake PowerPoint the way public/templates/bwb-slides.pptx lays them out.

import { vi } from 'vitest'
import type { FakePresentation, FakeSlide } from '../../fakePpt/index.ts'
import { line, placeholder, shape, slide, textBox } from '../../fakePpt/index.ts'

export const LIBRARY_FILE = 'templates/bwb-slides.pptx'
export const LIBRARY_BYTES = new TextEncoder().encode('bwb-slides.pptx stand-in')
export const LIBRARY_BASE64 = Buffer.from(LIBRARY_BYTES).toString('base64')

const AGENDA_ID = '257#136333637'
const DIVIDER_ID = '258#414069683'
const BULLETS_ID = '260#1940440457'

const entry = (id: string, title: string, category: string, slideId: string) => ({
  id, title, category, description: `${title}.`, tags: [id], slideId, thumb: `assets/slides/thumbs/${id}.png`,
})

export const LIBRARY_INDEX = {
  version: 1,
  file: LIBRARY_FILE,
  slides: [
    entry('agenda', 'Agenda', 'Structure', AGENDA_ID),
    entry('section-divider', 'Section divider', 'Structure', DIVIDER_ID),
    entry('bullets', 'Text and callout', 'Messages', BULLETS_ID),
  ],
}

/** Serves the library's index (or a 404 for null) and its .pptx. Returns the mock, to count the fetches. */
export function serveLibrary(index: object | null = LIBRARY_INDEX) {
  const fetchMock = vi.fn(async (url: string | URL | Request): Promise<Response> => {
    const u = String(url)
    if (u.endsWith('assets/slides/index.json')) {
      return index ? new Response(JSON.stringify(index), { status: 200 }) : new Response('Not found', { status: 404 })
    }
    if (u.endsWith(LIBRARY_FILE)) return new Response(LIBRARY_BYTES, { status: 200 })
    return new Response('Not found', { status: 404 })
  })
  vi.stubGlobal('fetch', fetchMock)
  return fetchMock
}

const slideNumber = (n: string) => placeholder('SlideNumber', n, { name: 'Slide number', left: 880, top: 508, width: 40, height: 22 })

/** The library's agenda: a title and five rows of number, item, description and rule, on the Title Only layout. */
export function agendaSlide(): FakeSlide {
  const rows = [1, 2, 3, 4, 5].flatMap((n) => {
    const top = 108 + (n - 1) * 72
    return [
      textBox(`0${n}`, { name: `Agenda number ${n}`, left: 40, top, width: 58.7, height: 72 }),
      textBox(`Sample item ${n}`, { name: `Agenda item ${n}`, left: 114.7, top, width: 357.3, height: 72 }),
      textBox(`Sample description ${n}`, { name: `Agenda description ${n}`, left: 488, top, width: 432, height: 72 }),
      line({ name: `Agenda rule ${n}`, left: 40, top: top + 72, width: 880, height: 0 }),
    ]
  })
  return slide([
    placeholder('Title', 'Agenda', { name: 'Title' }), line({ name: 'Agenda rule top', left: 40, top: 108, width: 880, height: 0 }), ...rows, slideNumber('2'),
  ], { layoutId: 'layout-title-only' })
}

/** The library's text slide: title, content placeholder, callout with its bar, and a source line. */
export function bulletsSlide(): FakeSlide {
  return slide([
    placeholder('Title', 'Action title: say the so-what in one sentence', { name: 'Title' }),
    placeholder('Content', 'Lead with the point\rThen the evidence', { name: 'Content', left: 40, top: 108, width: 581.3, height: 360 }),
    textBox('Key takeaway', { name: 'Callout', left: 637.3, top: 108, width: 282.7, height: 132 }),
    shape({ name: 'Callout bar', left: 637.3, top: 108, width: 4, height: 132 }),
    textBox('Source: organisation, title (year)', { name: 'Source', left: 40, top: 476, width: 700, height: 24, font: { size: 8 } }),
    slideNumber('5'),
  ])
}

/** A navy section divider on the Section Header layout. */
export function dividerSlide(title: string): FakeSlide {
  return slide([placeholder('Body', '01', { name: 'Section number', top: 120, height: 60 }), placeholder('Title', title, { name: 'Title' })], { layoutId: 'layout-section' })
}

/** Makes insertSlidesFromBase64 add fresh copies of the library slides it is asked for. */
export function useLibrarySlides(model: FakePresentation): void {
  const make: Record<string, () => FakeSlide> = {
    [AGENDA_ID]: agendaSlide, [DIVIDER_ID]: () => dividerSlide('Section title'), [BULLETS_ID]: bulletsSlide,
  }
  model.slidesFromBase64 = (_base64, options) => {
    const ids = Array.isArray(options.sourceSlideIds) ? options.sourceSlideIds.map(String) : []
    return ids.map((id) => {
      const build = make[id]
      if (!build) throw new Error(`No library slide ${id}`)
      return build()
    })
  }
}
