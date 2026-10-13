import { describe, expect, it } from 'vitest'
import { CONTENT_AREA, SAFE_AREA, SLIDE, columnEdges } from '../brand/slideGrid.ts'
import {
  areasFor, cmToPt, effectiveSize, gridColumns, largest, moveBoundsTo, normaliseRotation, ptToCm, resizeEffective, resolveReference,
  rightAngle, roundPt, setBoundsHeight, setBoundsWidth, translate, union, visualBounds,
} from './bounds.ts'
import type { Box } from './types.ts'

const box = (id: string, left: number, top: number, width: number, height: number, rotation?: number): Box =>
  rotation === undefined ? { id, left, top, width, height } : { id, left, top, width, height, rotation }

describe('units and rounding', () => {
  it('converts centimetres and points (1 cm = 72/2.54 pt)', () => {
    expect(cmToPt(2.54)).toBeCloseTo(72, 10)
    expect(ptToCm(72)).toBeCloseTo(2.54, 10)
    expect(cmToPt(0.2)).toBeCloseTo(5.669, 3)
  })

  it('rounds to a hundredth of a point without negative zero', () => {
    expect(roundPt(10.126)).toBe(10.13)
    expect(Object.is(roundPt(-0.001), 0)).toBe(true)
  })
})

describe('rotation', () => {
  it('normalises any angle to 0 up to 360', () => {
    expect(normaliseRotation(-90)).toBe(270)
    expect(normaliseRotation(450)).toBe(90)
    expect(normaliseRotation(360)).toBe(0)
    expect(normaliseRotation()).toBe(0)
  })

  it('recognises right angles to a hundredth of a degree', () => {
    expect(rightAngle(90.004)).toBe(90)
    expect(rightAngle(-90)).toBe(270)
    expect(rightAngle(359.999)).toBe(0)
    expect(rightAngle(30)).toBeNull()
  })
})

describe('visual bounds', () => {
  it('are the frame itself when unrotated or half turned', () => {
    expect(visualBounds(box('a', 10, 20, 100, 50))).toEqual({ left: 10, top: 20, width: 100, height: 50 })
    expect(visualBounds(box('a', 10, 20, 100, 50, 180))).toEqual({ left: 10, top: 20, width: 100, height: 50 })
  })

  it('swap width and height about the centre for a quarter turn', () => {
    expect(visualBounds(box('a', 0, 0, 100, 20, 90))).toEqual({ left: 40, top: -40, width: 20, height: 100 })
    expect(visualBounds(box('a', 0, 0, 100, 20, -90))).toEqual({ left: 40, top: -40, width: 20, height: 100 })
  })

  it('grow for other angles', () => {
    const v = visualBounds(box('a', 0, 0, 100, 100, 45))
    expect(v.width).toBeCloseTo(100 * Math.SQRT2, 6)
    expect(v.left).toBeCloseTo(50 - 50 * Math.SQRT2, 6)
    const w = visualBounds(box('b', 0, 0, 100, 20, 30))
    expect(w.width).toBeCloseTo(100 * Math.cos(Math.PI / 6) + 20 * 0.5, 6)
    expect(w.height).toBeCloseTo(100 * 0.5 + 20 * Math.cos(Math.PI / 6), 6)
  })

  it('unite into one rectangle, or none', () => {
    expect(union([])).toBeNull()
    expect(union([{ left: -10, top: 5, width: 20, height: 5 }, { left: 30, top: -5, width: 10, height: 30 }])).toEqual({ left: -10, top: -5, width: 50, height: 30 })
  })
})

describe('moving and resizing', () => {
  it('moves by rounding to hundredths, and leaves a zero move untouched', () => {
    const a = box('a', 10.004, 20, 100.123, 50)
    expect(translate(a, 0, 0)).toEqual(a)
    expect(translate(a, 1, 0)).toEqual({ ...a, left: 11, top: 20 })
  })

  it('moves the corner of the bounds to a point', () => {
    const r = moveBoundsTo(box('r', 0, 0, 100, 20, 90), 100, 100)
    expect(visualBounds(r)).toMatchObject({ left: 100, top: 100 })
    expect(moveBoundsTo(box('a', 5, 5, 10, 10), undefined, 50)).toMatchObject({ left: 5, top: 50 })
  })

  it('reads and sets the size as it reads on the slide, keeping the corner of the bounds', () => {
    const r = box('r', 0, 0, 100, 20, 90)
    expect(effectiveSize(r)).toEqual({ width: 20, height: 100 })
    const resized = resizeEffective(r, 30, 50)
    expect(resized).toMatchObject({ width: 50, height: 30, rotation: 90 })
    expect(visualBounds(resized)).toEqual({ left: 40, top: -40, width: 30, height: 50 })
    expect(resizeEffective(box('a', 5, 6, 10, 10), 20, 30)).toEqual(box('a', 5, 6, 20, 30))
  })

  it('sets the width of the bounds from either side', () => {
    expect(setBoundsWidth(box('a', 100, 0, 100, 50), 150, 'left')).toEqual(box('a', 100, 0, 150, 50))
    expect(setBoundsWidth(box('a', 100, 0, 100, 50), 150, 'right')).toEqual(box('a', 50, 0, 150, 50))
    expect(setBoundsWidth(box('a', 100, 0, 100, 50), 0, 'left')).toBeNull()
  })

  it('sets the width of rotated bounds through the frame side nearest horizontal', () => {
    const r = setBoundsWidth(box('r', 0, 0, 100, 20, 90), 50, 'left')
    expect(r).toMatchObject({ width: 100, height: 50 })
    expect(r && visualBounds(r)).toMatchObject({ left: 40, width: 50 })
    const tilted = box('t', 0, 0, 100, 20, 30)
    const before = visualBounds(tilted)
    const t = setBoundsWidth(tilted, 120, 'left')
    expect(t && visualBounds(t).width).toBeCloseTo(120, 1)
    expect(t && visualBounds(t).left).toBeCloseTo(before.left, 1)
    expect(t?.height).toBe(20)
    expect(setBoundsWidth(tilted, 5, 'left')).toBeNull()
  })

  it('sets the height of the bounds from either side', () => {
    expect(setBoundsHeight(box('a', 0, 100, 10, 50), 80, 'top')).toEqual(box('a', 0, 100, 10, 80))
    expect(setBoundsHeight(box('a', 0, 100, 10, 50), 80, 'bottom')).toEqual(box('a', 0, 70, 10, 80))
    const r = setBoundsHeight(box('r', 0, 0, 100, 20, 90), 60, 'top')
    expect(r && visualBounds(r)).toMatchObject({ top: -40, height: 60 })
  })
})

describe('slide areas and the grid', () => {
  it('match the BwB slide grid on a 16:9 slide', () => {
    const areas = areasFor()
    expect(areas.slide).toEqual(SLIDE)
    expect(areas.margins).toEqual(SAFE_AREA)
    expect(areas.content).toEqual(CONTENT_AREA)
    expect(gridColumns()).toEqual(columnEdges())
  })

  it('keep the margins on a slide of another size', () => {
    const areas = areasFor({ width: 720, height: 540 })
    expect(areas.margins).toEqual({ left: 40, top: 28, width: 640, height: 476 })
    expect(areas.content).toEqual({ left: 40, top: 108, width: 640, height: 360 })
    const cols = gridColumns({ width: 720, height: 540 })
    expect(cols[0]?.left).toBe(40)
    expect(cols[11]?.right).toBeCloseTo(680, 10)
  })
})

describe('references', () => {
  const a = box('a', 0, 0, 100, 50)
  const b = box('b', 200, 100, 50, 50)

  it('picks the largest box by the area of its bounds, the first of equals', () => {
    expect(largest([b, a])?.id).toBe('a')
    expect(largest([box('x', 0, 0, 10, 10), box('y', 0, 0, 10, 10)])?.id).toBe('x')
    expect(largest([])).toBeUndefined()
  })

  it('resolve to the selection, or the margins for one shape', () => {
    expect(resolveReference([a, b], { reference: 'selection' })).toEqual({ rect: { left: 0, top: 0, width: 250, height: 150 }, fixedId: null })
    expect(resolveReference([a], { reference: 'selection' }).rect).toEqual(SAFE_AREA)
  })

  it('resolve to the anchor, the largest by default, which stays put', () => {
    expect(resolveReference([a, b], { reference: 'anchor' })).toEqual({ rect: visualBounds(a), fixedId: 'a' })
    expect(resolveReference([a, b], { reference: 'anchor', anchor: b }).fixedId).toBe('b')
    const outside = box('z', 500, 500, 10, 10)
    expect(resolveReference([a, b], { reference: 'anchor', anchor: outside })).toEqual({ rect: visualBounds(outside), fixedId: 'z' })
    expect(resolveReference([], { reference: 'anchor' })).toEqual({ rect: SAFE_AREA, fixedId: null })
  })

  it('resolve to the slide, the margins and the content area', () => {
    expect(resolveReference([a], { reference: 'slide' }).rect).toEqual(SLIDE)
    expect(resolveReference([a], { reference: 'margins' }).rect).toEqual(SAFE_AREA)
    expect(resolveReference([a], { reference: 'content' }).rect).toEqual(CONTENT_AREA)
    expect(resolveReference([a], { reference: 'slide', slide: { width: 720, height: 540 } }).rect.width).toBe(720)
  })
})
