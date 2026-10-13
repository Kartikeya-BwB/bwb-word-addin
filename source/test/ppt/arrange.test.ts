import { afterEach, describe, expect, it } from 'vitest'
import { columnEdges } from '../../src/brand/slideGrid.ts'
import { cmToPt } from '../../src/geometry/index.ts'
import {
  alignShapes, distributeShapes, equaliseShapeGaps, gridShapes, groupShapes, matchShapeSize, nudgeShapes, orderShapes, readShapeFrame,
  setShapeFrame, snapShapes, stackShapes, stretchShapes, swapShapes, tidyShapes, ungroupShapes,
} from '../../src/ppt/arrange.ts'
import { selectionTracker } from '../../src/ppt/selection.ts'
import type { FakePresentation, FakeShape } from '../fakePpt/index.ts'
import { clickShapes, createPresentation, group, installFakePowerPoint, picture, select, shape, slide, uninstallFakePowerPoint } from '../fakePpt/index.ts'

afterEach(async () => {
  await selectionTracker.stop()
  selectionTracker.clearAnchor()
  selectionTracker.apply({ slideId: null, shapes: [] })
  uninstallFakePowerPoint()
})

function deck(shapes: FakeShape[], selected: string[], over: Partial<FakePresentation> = {}): FakePresentation {
  const d = createPresentation([slide(shapes, { id: 's1' })], over)
  select(d, selected)
  return d
}

const abc = (): FakeShape[] => [
  shape({ id: 'a', left: 0, top: 0, width: 100, height: 50 }),
  shape({ id: 'b', left: 200, top: 100, width: 50, height: 50 }),
  shape({ id: 'c', left: 50, top: 300, width: 80, height: 40 }),
]

const lefts = (shapes: FakeShape[]): number[] => shapes.map((s) => s.left)

describe('reading and writing back', () => {
  it('lines shapes up and writes the changes in one sync', async () => {
    const shapes = abc()
    const fake = installFakePowerPoint(deck(shapes, ['a', 'b', 'c']))
    expect(await alignShapes('left', 'selection')).toBe('Lined up 2 shapes.')
    expect(lefts(shapes)).toEqual([0, 0, 0])
    expect(fake.runs()).toBe(1)
    expect(fake.syncs()).toBe(3)
  })

  it('says so when nothing needs to move, without writing', async () => {
    const shapes = abc().map((s) => ({ ...s, left: 10 }))
    const fake = installFakePowerPoint(deck(shapes, ['a', 'b', 'c']))
    expect(await alignShapes('left', 'selection')).toBe('Nothing to change: the shapes are already in place.')
    expect(fake.syncs()).toBe(2)
  })

  it('lines a single shape up with the margins', async () => {
    const shapes = abc()
    installFakePowerPoint(deck(shapes, ['b']))
    await alignShapes('right', 'selection')
    expect(shapes[1]?.left).toBe(870)
  })

  it('moves the selection as a group to the slide, the margins and the content area', async () => {
    const shapes = abc()
    installFakePowerPoint(deck(shapes, ['a', 'b']))
    await alignShapes('centre', 'slide', true)
    expect(lefts(shapes)).toEqual([355, 555, 50])
    await alignShapes('top', 'content', true)
    expect(shapes.map((s) => s.top)).toEqual([108, 208, 300])
  })

  it('uses the slide size PowerPoint reports', async () => {
    const shapes = [shape({ id: 'a', width: 100 })]
    installFakePowerPoint(deck(shapes, ['a'], { pageSetup: { slideWidth: 720, slideHeight: 540 } }))
    await alignShapes('centre', 'slide')
    expect(shapes[0]?.left).toBe(310)
  })

  it('lines rotated shapes up by their bounds, and treats them as upright where PowerPoint cannot say', async () => {
    const r = shape({ id: 'r', left: 0, top: 0, width: 100, height: 20, rotation: 90 })
    const a = shape({ id: 'a', left: 200, top: 0, width: 50, height: 50 })
    installFakePowerPoint(deck([r, a], ['r', 'a']))
    await alignShapes('left', 'selection')
    expect([r.left, a.left]).toEqual([0, 40])
    const r2 = shape({ id: 'r', left: 0, top: 0, width: 100, height: 20, rotation: 90 })
    const a2 = shape({ id: 'a', left: 200, top: 0, width: 50, height: 50 })
    installFakePowerPoint(deck([r2, a2], ['r', 'a'], { apiVersion: '1.8' }))
    await alignShapes('left', 'selection')
    expect([r2.left, a2.left]).toEqual([0, 0])
  })

  it('moves and scales groups through their members', async () => {
    const g = group([shape({ id: 'm1', left: 100, top: 100, width: 50, height: 50 }), shape({ id: 'm2', left: 200, top: 100, width: 50, height: 50 })], { id: 'g' })
    installFakePowerPoint(deck([g], ['g']))
    await alignShapes('left', 'margins')
    expect(g.members?.map((m) => m.left)).toEqual([40, 140])
  })

  it('needs PowerPoint, and a PowerPoint new enough to report the selection', async () => {
    await expect(alignShapes('left', 'selection')).rejects.toThrow(/Open this in PowerPoint/)
    installFakePowerPoint(deck(abc(), ['a']), { apiVersion: '1.4' })
    await expect(alignShapes('left', 'selection')).rejects.toThrow('Arranging shapes needs a newer PowerPoint. Update Microsoft 365, or use PowerPoint on the web.')
  })

  it('asks for a selection', async () => {
    installFakePowerPoint(deck(abc(), []))
    await expect(alignShapes('left', 'selection')).rejects.toThrow('Select the shapes first.')
  })
})

describe('the anchor', () => {
  it('is the first shape selected, when the pane saw the selection being made', async () => {
    const shapes = abc()
    const d = deck(shapes, [])
    installFakePowerPoint(d)
    await selectionTracker.start()
    await clickShapes(d, ['b'])
    await clickShapes(d, ['a', 'b'])
    await clickShapes(d, ['a', 'b', 'c'])
    await matchShapeSize('width', 'anchor')
    expect(shapes.map((s) => s.width)).toEqual([50, 50, 50])
  })

  it('is the largest shape when the order of selection is unknown', async () => {
    const shapes = abc()
    installFakePowerPoint(deck(shapes, ['a', 'b', 'c']))
    await matchShapeSize('height', 'anchor')
    expect(shapes.map((s) => s.height)).toEqual([50, 50, 50])
    await alignShapes('right', 'anchor')
    expect(lefts(shapes)).toEqual([0, 50, 20])
  })

  it('can be set by hand and then works from outside the selection', async () => {
    const shapes = [...abc(), shape({ id: 'z', name: 'Key box', left: 600, top: 400, width: 77, height: 33 })]
    const d = deck(shapes, [])
    installFakePowerPoint(d)
    await clickShapes(d, ['z'])
    await selectionTracker.setAnchorFromSelection()
    select(d, ['a', 'b'])
    expect(await matchShapeSize('both', 'anchor')).toBe('Resized 2 shapes.')
    expect(shapes.map((s) => [s.width, s.height])).toEqual([[77, 33], [77, 33], [80, 40], [77, 33]])
    await alignShapes('bottom', 'anchor')
    expect(shapes.map((s) => s.top)).toEqual([400, 400, 300, 400])
  })

  it('counts an anchor outside the selection only for the tools that use it', async () => {
    const shapes = [...abc(), shape({ id: 'z', left: 600, top: 400, width: 77, height: 33 })]
    const d = deck(shapes, ['z'])
    installFakePowerPoint(d)
    await selectionTracker.setAnchorFromSelection()
    select(d, ['a'])
    expect(await matchShapeSize('width', 'anchor')).toBe('Resized 1 shape.')
    await expect(stackShapes('horizontal')).rejects.toThrow('Select two or more shapes to stack.')
    await expect(swapShapes()).rejects.toThrow('Select two shapes to swap.')
  })

  it('needs two shapes in all to share a size', async () => {
    installFakePowerPoint(deck(abc(), ['a']))
    await expect(matchShapeSize('width', 'anchor')).rejects.toThrow('Select two or more shapes. The anchor, or the first one you select, sets the size.')
  })
})

describe('the other arrange tools', () => {
  it('spreads shapes out, and needs three between the outer two', async () => {
    const shapes = [shape({ id: 'a', left: 0, width: 100 }), shape({ id: 'b', left: 150, width: 50 }), shape({ id: 'c', left: 400, width: 100 })]
    installFakePowerPoint(deck(shapes, ['a', 'b', 'c']))
    await distributeShapes({ axis: 'horizontal' })
    expect(lefts(shapes)).toEqual([0, 225, 400])
    await distributeShapes({ axis: 'horizontal', within: 'margins' })
    expect(lefts(shapes)).toEqual([40, 455, 820])
    await distributeShapes({ axis: 'horizontal', gap: cmToPt(1) })
    expect(shapes[1]?.left).toBeCloseTo(40 + 100 + 28.35, 1)
    installFakePowerPoint(deck(abc(), ['a', 'b']))
    await expect(distributeShapes({ axis: 'vertical' })).rejects.toThrow('Select three or more shapes to space them out between the outer two.')
  })

  it('evens out gaps, stacks, tiles, tidies and snaps', async () => {
    const shapes = [shape({ id: 'a', left: 0, top: 0, width: 100 }), shape({ id: 'b', left: 150, top: 3, width: 50 }), shape({ id: 'c', left: 210, top: 0, width: 100 })]
    installFakePowerPoint(deck(shapes, ['a', 'b', 'c']))
    await equaliseShapeGaps('horizontal')
    expect(lefts(shapes)).toEqual([0, 130, 210])
    await stackShapes('horizontal')
    expect(lefts(shapes)).toEqual([0, 100, 150])
    await gridShapes({ columns: 2, gapX: 10, gapY: 10 })
    expect(shapes.map((s) => [s.left, s.top])).toEqual([[0, 0], [110, 0], [0, 60]])
    await tidyShapes()
    expect(shapes[2]?.left).toBe(0)
    await snapShapes()
    expect(shapes[0]?.left).toBeCloseTo(columnEdges()[0]?.left ?? NaN, 2)
  })

  it('swaps two shapes and refuses any other number', async () => {
    const shapes = abc()
    installFakePowerPoint(deck(shapes, ['a', 'b']))
    expect(await swapShapes()).toBe('Swapped the two shapes.')
    expect(shapes.slice(0, 2).map((s) => [s.left, s.top])).toEqual([[200, 100], [0, 0]])
    installFakePowerPoint(deck(abc(), ['a', 'b', 'c']))
    await expect(swapShapes()).rejects.toThrow('Select two shapes to swap.')
    installFakePowerPoint(deck(abc(), ['a']))
    await expect(swapShapes()).rejects.toThrow('Select two shapes to swap.')
  })

  it('stretches to a reference, nudges and sets exact frames', async () => {
    const shapes = abc()
    installFakePowerPoint(deck(shapes, ['a']))
    await stretchShapes('right', 'margins')
    expect(shapes[0]?.width).toBe(920)
    await nudgeShapes(cmToPt(0.1), 0)
    expect(shapes[0]?.left).toBeCloseTo(2.83, 2)
    await setShapeFrame({ left: cmToPt(2), width: cmToPt(5) })
    expect(shapes[0]?.left).toBeCloseTo(56.69, 2)
    expect(shapes[0]?.width).toBeCloseTo(141.73, 2)
    await expect(setShapeFrame({ width: 0 })).rejects.toThrow('Width and height must be more than zero.')
  })

  it('reads the frame of the anchor or the first shape, and nothing without a selection', async () => {
    installFakePowerPoint(deck(abc(), ['a', 'b']))
    expect(await readShapeFrame()).toEqual({ left: 0, top: 0, width: 100, height: 50 })
    installFakePowerPoint(deck(abc(), []))
    expect(await readShapeFrame()).toBeNull()
  })

  it('leaves pictures and lines to be moved like any shape', async () => {
    const p = picture({ id: 'p', left: 300, top: 300, width: 40, height: 40 })
    installFakePowerPoint(deck([p], ['p']))
    await alignShapes('top', 'content')
    expect(p.top).toBe(108)
  })
})

describe('order and groups', () => {
  const ids = (d: FakePresentation): string[] => d.slides[0]?.shapes.map((s) => s.id) ?? []

  it('brings shapes to the front and sends them to the back, keeping their own order', async () => {
    const d = deck([shape({ id: 'a' }), shape({ id: 'b' }), shape({ id: 'c' }), shape({ id: 'd' })], ['a', 'b'])
    installFakePowerPoint(d)
    await orderShapes('BringToFront')
    expect(ids(d)).toEqual(['c', 'd', 'a', 'b'])
    await orderShapes('SendToBack')
    expect(ids(d)).toEqual(['a', 'b', 'c', 'd'])
    await orderShapes('BringForward')
    expect(ids(d)).toEqual(['c', 'a', 'b', 'd'])
    expect(await orderShapes('SendBackward')).toBe('Sent backward: 2 shapes.')
    expect(ids(d)).toEqual(['a', 'b', 'c', 'd'])
  })

  it('groups and ungroups', async () => {
    const d = deck([shape({ id: 'a' }), shape({ id: 'b' }), shape({ id: 'c' })], ['a', 'b'])
    installFakePowerPoint(d)
    expect(await groupShapes()).toBe('Grouped 2 shapes.')
    expect(d.slides[0]?.shapes.map((s) => s.type)).toEqual(['Group', 'GeometricShape'])
    expect(await ungroupShapes()).toBe('Ungrouped 1 group.')
    expect(ids(d)).toEqual(['a', 'b', 'c'])
    await expect(ungroupShapes()).rejects.toThrow('Select a group to ungroup.')
    select(d, ['c'])
    await expect(groupShapes()).rejects.toThrow('Select two or more shapes to group.')
  })

  it('says plainly when PowerPoint is too old to order or group', async () => {
    installFakePowerPoint(deck(abc(), ['a', 'b']), { apiVersion: '1.7' })
    await expect(orderShapes('BringToFront')).rejects.toThrow('Changing the order of shapes needs a newer PowerPoint.')
    await expect(groupShapes()).rejects.toThrow('Grouping shapes needs a newer PowerPoint.')
    await expect(ungroupShapes()).rejects.toThrow('Ungrouping shapes needs a newer PowerPoint.')
  })
})
