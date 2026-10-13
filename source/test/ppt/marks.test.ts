import { afterEach, describe, expect, it, vi } from 'vitest'
import { ICON_SIZES_CM } from '../../src/brand/icons.ts'
import { CONTENT_AREA, TAGS } from '../../src/brand/slideGrid.ts'
import { cmToPt } from '../../src/geometry/index.ts'
import {
  addReviewNote, formatDate, insertHarveyBall, insertRagDot, insertStatusMark, isoDate, removeAllNotes, removeAllStickers, setSlideStickers,
} from '../../src/ppt/marks.ts'
import { harveySvg, statusSvg } from '../../src/ppt/svg.ts'
import type { FakePresentation } from '../fakePpt/index.ts'
import { createPresentation, installFakePowerPoint, select, shape, shapesTagged, slide, textOf, uninstallFakePowerPoint } from '../fakePpt/index.ts'
import { iconFetch, iconIndex } from './helpers.ts'

afterEach(() => {
  uninstallFakePowerPoint()
  vi.unstubAllGlobals()
})

function deck(apiVersion = '1.10'): FakePresentation {
  const d = createPresentation([slide([shape({ id: 'x' })], { id: 's1' }), slide([], { id: 's2' }), slide([], { id: 's3' })], { apiVersion })
  installFakePowerPoint(d)
  return d
}

describe('dates', () => {
  it('writes dates the BwB way, and as ISO for tags', () => {
    const d = new Date(2026, 9, 3)
    expect(formatDate(d)).toBe('3 October 2026')
    expect(isoDate(d)).toBe('2026-10-03')
  })
})

describe('stickers', () => {
  it('puts a sticker on each selected slide and replaces one already there', async () => {
    const d = deck()
    d.selection.slideIds = ['s1', 's3']
    expect(await setSlideStickers('Draft')).toBe('“Draft” sticker on 2 slides.')
    expect(await setSlideStickers('For discussion')).toBe('“For discussion” sticker on 2 slides.')
    const stickers = shapesTagged(d, TAGS.sticker)
    expect(stickers.map((s) => [textOf(s), s.tags.get(TAGS.sticker), s.tags.get(TAGS.element)])).toEqual([
      ['FOR DISCUSSION', 'For discussion', 'sticker'], ['FOR DISCUSSION', 'For discussion', 'sticker'],
    ])
    expect(shapesTagged(d.slides[1] ?? slide(), TAGS.sticker)).toHaveLength(0)
    const [first] = stickers
    expect(first?.top).toBeLessThan(28)
    expect((first?.left ?? 0) + (first?.width ?? 0)).toBeCloseTo(920, 6)
    expect(first?.textFrame?.runs[0]?.font).toMatchObject({ name: 'Aptos', size: 11, bold: true, color: '#4A4A4C' })
  })

  it('removes every sticker in the deck and leaves other shapes alone', async () => {
    const d = deck()
    d.selection.slideIds = ['s1', 's2', 's3']
    await setSlideStickers('Confidential')
    expect(await removeAllStickers()).toBe('Removed 3 stickers.')
    expect(shapesTagged(d, TAGS.sticker)).toHaveLength(0)
    expect(d.slides[0]?.shapes.map((s) => s.id)).toEqual(['x'])
    expect(await removeAllStickers()).toBe('There are no stickers in this presentation.')
  })
})

describe('review notes', () => {
  it('adds a dated navy note, nudging each new one past the last, and removes them all', async () => {
    const d = deck()
    const today = new Date(2026, 9, 13)
    await addReviewNote(today)
    await addReviewNote(today)
    const notes = shapesTagged(d, TAGS.note)
    expect(notes.map((n) => [textOf(n), n.tags.get(TAGS.note), n.fill.foregroundColor])).toEqual([
      ['Note: \n13 October 2026', '2026-10-13', '#022338'], ['Note: \n13 October 2026', '2026-10-13', '#022338'],
    ])
    expect((notes[1]?.top ?? 0) - (notes[0]?.top ?? 0)).toBe(8)
    expect(notes[0]?.textFrame?.runs.every((r) => r.font.color === '#FFFFFF')).toBe(true)
    expect(await removeAllNotes()).toBe('Removed 2 review notes.')
    expect(await removeAllNotes()).toBe('There are no review notes in this presentation.')
  })
})

describe('Harvey balls', () => {
  it('draws a navy ring filled by quarters', () => {
    expect(harveySvg(0)).not.toContain('<path')
    expect(harveySvg(2)).toContain('A9 9 0 0 1 12 21')
    expect(harveySvg(4)).toContain('<circle cx="12" cy="12" r="9" fill="#022338"/>')
    expect(harveySvg(3)).toContain('stroke="#022338"')
  })

  it('inserts SVG in the content area, names and tags it, and puts the next one alongside', async () => {
    const d = deck()
    await insertHarveyBall(3)
    await insertHarveyBall(1)
    const size = cmToPt(ICON_SIZES_CM.marker)
    expect(d.dataInserts.map((i) => i.coercionType)).toEqual(['xmlSvg', 'xmlSvg'])
    expect(d.dataInserts[0]?.data).toBe(harveySvg(3))
    const balls = shapesTagged(d, TAGS.element, 'harvey')
    expect(balls.map((b) => [b.type, b.name])).toEqual([['Graphic', 'BwB Harvey ball 3'], ['Graphic', 'BwB Harvey ball 1']])
    expect(balls[0]).toMatchObject({ left: CONTENT_AREA.left + (CONTENT_AREA.width - size) / 2, width: size })
    expect(balls[1]?.left).toBeCloseTo((balls[0]?.left ?? 0) + size + 8, 6)
    expect(balls[0]?.altTextDescription).toBe('Harvey ball, three quarters (75 per cent)')
  })

  it('leaves the graphic untagged when PowerPoint selects something else after inserting', async () => {
    const d = deck()
    type Insert = (data: string, options: Office.SetSelectedDataOptions, cb: (r: { status: string }) => void) => void
    const original = Office.context.document.setSelectedDataAsync as unknown as Insert
    const insert: Insert = (data, options, cb) => {
      original.call(Office.context.document, data, options, () => {
        select(d, ['x'], 's1')
        cb({ status: 'succeeded' })
      })
    }
    Object.assign(Office.context.document, { setSelectedDataAsync: insert })
    await insertHarveyBall(2)
    expect(shapesTagged(d, TAGS.element)).toHaveLength(0)
  })

  it('reports a failed insert plainly', async () => {
    const d = deck()
    d.failNextDataInsert = 'The slide is locked'
    await expect(insertHarveyBall(0)).rejects.toThrow('PowerPoint did not insert the graphic (The slide is locked).')
  })
})

describe('status marks and RAG dots', () => {
  it('builds Lucide status marks in BwB colours', () => {
    expect(statusSvg(iconIndex(), 'onTrack')).toContain('stroke="#328840"')
    expect(statusSvg(iconIndex(), 'atRisk')).toContain('stroke="#4A4A4C"')
    expect(statusSvg(iconIndex(), 'offTrack')).toContain('m15 9-6 6')
  })

  it('inserts a status mark from the icon library the pane serves', async () => {
    vi.stubGlobal('fetch', iconFetch)
    const d = deck()
    expect(await insertStatusMark('onTrack')).toBe('“On track” mark added.')
    expect(await insertStatusMark('offTrack', async () => iconIndex())).toBe('“Off track” mark added.')
    expect(shapesTagged(d, TAGS.element, 'status').map((s) => s.name)).toEqual(['BwB status on track', 'BwB status off track'])
  })

  it('adds RAG dots as small ellipses tagged with their colour, in a row', async () => {
    const d = deck()
    expect(await insertRagDot('green')).toBe('Green dot added.')
    expect(await insertRagDot('amber')).toMatch(/^Amber dot added. It is outside the BwB palette/)
    await insertRagDot('red')
    const dots = shapesTagged(d, TAGS.rag)
    const size = cmToPt(ICON_SIZES_CM.inline)
    expect(dots.map((s) => [s.geometricShapeType, s.fill.foregroundColor, s.tags.get(TAGS.rag), s.tags.get(TAGS.element), s.line.visible])).toEqual([
      ['Ellipse', '#328840', 'green', 'rag', false], ['Ellipse', '#F2A900', 'amber', 'rag', false], ['Ellipse', '#A72C28', 'red', 'rag', false],
    ])
    expect(dots[0]?.width).toBeCloseTo(size, 6)
    expect((dots[1]?.left ?? 0) - (dots[0]?.left ?? 0)).toBeCloseTo(size + 8, 1)
    expect(d.selection.shapeIds).toEqual([dots[2]?.id])
  })

  it('needs a PowerPoint that reports the selection', async () => {
    deck('1.4')
    await expect(insertRagDot('green')).rejects.toThrow('Inserting marks needs a newer PowerPoint.')
    await expect(setSlideStickers('Draft')).rejects.toThrow('Adding stickers needs a newer PowerPoint.')
  })
})
