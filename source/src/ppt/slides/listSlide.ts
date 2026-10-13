// A new slide after the current one with a title and a list: the agenda and the executive summary. It starts from
// the BwB library slide when there is one, or else a plain slide with a title and content layout. The BwB agenda
// slide's numbered rows are filled in place; otherwise the list goes into the slide's body placeholder (or a new text
// box on the BwB grid), and the library's sample shapes the slide no longer needs are removed.

import { CONTENT_AREA, SAFE_AREA, SLIDE_TYPE, TAGS, TITLE_AREA } from '../../brand/slideGrid.ts'
import { COLOURS } from '../../brand/tokens.ts'
import { UserFacingError } from '../../word/run.ts'
import { BODY_PLACEHOLDERS, isBwbFurniture, isFooterShape, isTitleShape, tagValue } from '../check/shapes.ts'
import { takeDeckSnapshot } from '../check/snapshot.ts'
import type { ShapeInfo, SlideInfo } from '../check/types.ts'
import type { InsertedSlides, LibrarySlide, SlideLibraryIndex } from './library.ts'
import { findLibrarySlide, insertLibrarySlides, loadSlideIndex } from './library.ts'
import { NEEDS, shapeAt } from './office.ts'
import { requirePpt, runPpt, supportsPpt } from '../run.ts'

export interface ListSlideRequest {
  /** Library slides to start from, by id or category, the first found wins ('agenda'; 'content', 'bullets'). */
  library: string[]
  title: string
  items: string[]
  /** Numbered for an agenda; bullets for a summary. */
  numbered: boolean
  /** Written to the slide's BWB_LAYOUT tag, so the slide tools know what it is. */
  layout: string
}

export type FillStep =
  | { kind: 'text'; path: string[]; text: string; bullets?: boolean }
  | { kind: 'remove'; path: string[] }
  | { kind: 'resize'; path: string[]; left: number; width: number }
  | { kind: 'add-title'; text: string }
  | { kind: 'add-list'; text: string; bullets: boolean }

const LIST_ELEMENT = /list|agenda|body|bullet|content/i
/** The BwB agenda slide's rows: "Agenda number 1", "Agenda item 1", "Agenda description 1", "Agenda rule 1". */
const AGENDA_ROW = /^agenda (number|item|description|rule) (\d+)$/i
const RIGHT_EDGE = SAFE_AREA.left + SAFE_AREA.width

const area = (s: ShapeInfo): number => s.width * s.height

/** Where the title and the list go: its placeholders, a shape tagged for it, or the largest text shape below the title. */
export function listSlidePlan(slide: SlideInfo): { title: ShapeInfo | null; list: ShapeInfo | null } {
  const usable = slide.shapes.filter((s) => s.level === 0 && s.text !== null && !isBwbFurniture(s) && !isFooterShape(s))
  const title = usable.find(isTitleShape) ?? usable.find((s) => /title/i.test(tagValue(s, TAGS.element) ?? '')) ?? null
  const others = usable.filter((s) => s !== title)
  const list =
    others.find((s) => LIST_ELEMENT.test(tagValue(s, TAGS.element) ?? '')) ??
    others.find((s) => s.placeholder !== null && BODY_PLACEHOLDERS.has(s.placeholder)) ??
    others.filter((s) => s.top >= TITLE_AREA.top + TITLE_AREA.height - 8).sort((a, b) => area(b) - area(a))[0] ??
    null
  return { title, list }
}

function agendaRows(slide: SlideInfo): Map<number, Map<string, ShapeInfo>> {
  const rows = new Map<number, Map<string, ShapeInfo>>()
  for (const s of slide.shapes) {
    const m = s.level === 0 ? AGENDA_ROW.exec(s.name.trim()) : null
    if (!m?.[1] || !m[2]) continue
    const n = Number(m[2])
    const row = rows.get(n) ?? new Map<string, ShapeInfo>()
    row.set(m[1].toLowerCase(), s)
    rows.set(n, row)
  }
  return rows
}

/** The steps that fill a new slide with the title and items. */
export function planListSlide(slide: SlideInfo, request: Pick<ListSlideRequest, 'title' | 'items' | 'numbered'>): FillStep[] {
  const rows = agendaRows(slide)
  // Row shapes are never the list: when the rows are not used, they go.
  const inRows = new Set([...rows.values()].flatMap((row) => [...row.values()]))
  const { title, list } = listSlidePlan({ ...slide, shapes: slide.shapes.filter((s) => !inRows.has(s)) })
  const steps: FillStep[] = [title ? { kind: 'text', path: title.path, text: request.title } : { kind: 'add-title', text: request.title }]

  const count = request.items.length
  const fits = count > 0 && count <= rows.size && request.items.every((_, i) => rows.get(i + 1)?.has('item'))
  if (request.numbered && fits) {
    for (const [n, row] of [...rows].sort((a, b) => a[0] - b[0])) {
      const item = request.items[n - 1]
      if (item === undefined) {
        row.forEach((s) => steps.push({ kind: 'remove', path: s.path }))
        continue
      }
      const number = row.get('number')
      const label = row.get('item')
      const description = row.get('description')
      if (number) steps.push({ kind: 'text', path: number.path, text: String(n).padStart(2, '0') })
      if (label) {
        steps.push({ kind: 'text', path: label.path, text: item })
        // The library's descriptions are sample text: the section names take the whole row instead.
        if (description) steps.push({ kind: 'resize', path: label.path, left: label.left, width: RIGHT_EDGE - label.left })
      }
      if (description) steps.push({ kind: 'remove', path: description.path })
    }
    return steps
  }

  const text = request.items.map((item, i) => (request.numbered ? `${i + 1}. ${item}` : item)).join('\n')
  const keep = new Set<ShapeInfo>([title, list].filter((s): s is ShapeInfo => s !== null))
  if (list) {
    steps.push({ kind: 'text', path: list.path, text, bullets: !request.numbered })
    if (Math.abs(list.left - CONTENT_AREA.left) <= 2 && list.width < CONTENT_AREA.width - 2) {
      steps.push({ kind: 'resize', path: list.path, left: CONTENT_AREA.left, width: CONTENT_AREA.width })
    }
  } else {
    steps.push({ kind: 'add-list', text, bullets: !request.numbered })
  }
  // Whatever else came with a library slide is its sample content: callouts, sources, columns.
  for (const s of slide.shapes) {
    if (s.level === 0 && !keep.has(s) && !isFooterShape(s) && !isBwbFurniture(s)) steps.push({ kind: 'remove', path: s.path })
  }
  return steps
}

/** The layout for a plain list slide: title and content, else title only, from the layouts of the current master. */
export function pickLayout(layouts: Array<{ id: string; name: string }>): { id: string; name: string } | undefined {
  return layouts.find((l) => /title,? (?:and|&) content/i.test(l.name)) ?? layouts.find((l) => /title only/i.test(l.name))
}

/** The first library slide the request names. */
export function chooseLibrarySlide(index: SlideLibraryIndex, keys: string[]): LibrarySlide | undefined {
  for (const key of keys) {
    const found = findLibrarySlide(index, key)
    if (found) return found
  }
  return undefined
}

/** Adds a plain slide after the current one (or at the end), selects it and returns its id. */
async function addPlainSlide(): Promise<InsertedSlides> {
  const current = supportsPpt(NEEDS.current)
  const canMove = supportsPpt(NEEDS.move)
  return runPpt(async (context) => {
    const presentation = context.presentation
    const before = presentation.slides
    before.load('items/id')
    const selected = current ? presentation.getSelectedSlides() : null
    selected?.load('items/id')
    await context.sync()
    // presentation.slides is one proxy: read what it lists now, before the reload below lists the new slide too.
    const known = new Set(before.items.map((s) => s.id))
    const count = before.items.length
    const active = selected?.items[0]
    const activeIndex = active ? before.items.findIndex((s) => s.id === active.id) : -1
    const anchor = active ?? before.items[count - 1]
    let options: PowerPoint.AddSlideOptions = {}
    if (anchor) {
      const master = anchor.slideMaster
      master.load('id')
      const layouts = master.layouts
      layouts.load('items/id,items/name')
      await context.sync()
      const layout = pickLayout(layouts.items.map((l) => ({ id: l.id, name: l.name })))
      options = { slideMasterId: master.id, ...(layout ? { layoutId: layout.id } : {}) }
    }
    presentation.slides.add(options)
    await context.sync()
    const after = presentation.slides
    after.load('items/id')
    await context.sync()
    const added = after.items.find((s) => !known.has(s.id))
    if (!added) throw new UserFacingError('PowerPoint did not add the slide. Try again.')
    let afterCurrent = activeIndex >= 0 && activeIndex === count - 1
    if (activeIndex >= 0 && !afterCurrent && canMove) {
      added.moveTo(activeIndex + 1)
      afterCurrent = true
    }
    if (current) presentation.setSelectedSlides([added.id])
    await context.sync()
    return { slideIds: [added.id], afterCurrent }
  })
}

function styleText(range: PowerPoint.TextRange, spec: { font: string; size: number; bold: boolean; colour: string }): void {
  range.font.name = spec.font
  range.font.size = spec.size
  range.font.bold = spec.bold
  range.font.color = `#${spec.colour}`
}

/** Writes the title and list into the new slide, removes the sample shapes it does not need, and tags it. */
async function fillListSlide(slideId: string, request: ListSlideRequest): Promise<void> {
  const snapshot = await takeDeckSnapshot({ formatting: false, slideIds: [slideId] })
  const slide = snapshot.slides[0]
  if (!slide) throw new UserFacingError('PowerPoint did not add the slide. Try again.')
  const steps = planListSlide(slide, request)
  await runPpt(async (context) => {
    const target = context.presentation.slides.getItem(slideId)
    for (const step of steps) {
      switch (step.kind) {
        case 'text': {
          const range = shapeAt(target, step.path).textFrame.textRange
          range.text = step.text
          if (step.bullets !== undefined) range.paragraphFormat.bulletFormat.visible = step.bullets
          break
        }
        case 'resize': {
          const shape = shapeAt(target, step.path)
          shape.left = step.left
          shape.width = step.width
          break
        }
        case 'remove':
          shapeAt(target, step.path).delete()
          break
        case 'add-title': {
          const box = target.shapes.addTextBox(step.text, { ...TITLE_AREA })
          box.name = 'Title'
          styleText(box.textFrame.textRange, { ...SLIDE_TYPE.title })
          break
        }
        case 'add-list': {
          const box = target.shapes.addTextBox(step.text, { ...CONTENT_AREA })
          box.name = step.bullets ? 'Bullets' : 'List'
          styleText(box.textFrame.textRange, { ...SLIDE_TYPE.body, colour: COLOURS.black })
          box.textFrame.textRange.paragraphFormat.bulletFormat.visible = step.bullets
          break
        }
      }
    }
    target.tags.add(TAGS.layout, request.layout)
    await context.sync()
  })
}

/** Inserts a title-and-list slide after the current slide. */
export async function insertListSlide(request: ListSlideRequest): Promise<InsertedSlides> {
  requirePpt(NEEDS.read, `Inserting the ${request.title.toLowerCase()} slide`)
  const index = await loadSlideIndex().catch(() => null)
  const entry = index ? chooseLibrarySlide(index, request.library) : undefined
  const inserted = index && entry ? await insertLibrarySlides(index, [entry]) : await addPlainSlide()
  const first = inserted.slideIds[0]
  if (!first) throw new UserFacingError('PowerPoint did not add the slide. Try again.')
  await fillListSlide(first, request)
  return inserted
}
