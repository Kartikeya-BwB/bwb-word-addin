import { afterEach, describe, expect, it } from 'vitest'
import {
  boundsOf, clickShapes, createPresentation, fireSelectionChanged, group, installFakePowerPoint, picture, placeholder, select, selectText,
  shape, slide, textBox, textOf, uninstallFakePowerPoint,
} from './index.ts'

afterEach(uninstallFakePowerPoint)

const ppt = <T>(fn: (context: PowerPoint.RequestContext) => Promise<T>): Promise<T> => PowerPoint.run(fn)
describe('load and sync', () => {
  it('reads a property only after load() and sync()', async () => {
    const deck = createPresentation([slide([shape({ id: 'a', left: 10 })], { id: 's1' })])
    select(deck, ['a'])
    installFakePowerPoint(deck)
    await ppt(async (context) => {
      const shapes = context.presentation.getSelectedShapes()
      shapes.load('items/left')
      expect(() => shapes.items).toThrow(/read before load/)
      await context.sync()
      const first = shapes.items[0]
      expect(first?.left).toBe(10)
      expect(() => first?.top).toThrow(/Shape.*top read before load/)
    })
  })

  it('queues writes until sync, then applies them in order', async () => {
    const a = shape({ id: 'a', left: 10 })
    const deck = createPresentation([slide([a])])
    select(deck, ['a'])
    const fake = installFakePowerPoint(deck)
    await ppt(async (context) => {
      const s = context.presentation.getSelectedShapes().getItem('a')
      s.left = 20
      s.left = 30
      expect(a.left).toBe(10)
      await context.sync()
      expect(a.left).toBe(30)
    })
    expect(fake.runs()).toBe(1)
    expect(fake.syncs()).toBe(2)
  })

  it('loads nested paths, option objects and the scalars of a navigation property', async () => {
    const a = shape({ id: 'a', text: 'Hello', font: { bold: true } })
    const deck = createPresentation([slide([a])])
    select(deck, ['a'])
    installFakePowerPoint(deck)
    await ppt(async (context) => {
      const shapes = context.presentation.getSelectedShapes()
      shapes.load('items/textFrame/textRange/text,items/fill')
      const one = shapes.getItem('a')
      one.load({ textFrame: { textRange: { font: { bold: true } } }, name: true })
      await context.sync()
      expect(shapes.items[0]?.textFrame.textRange.text).toBe('Hello')
      expect(shapes.items[0]?.fill.foregroundColor).toBe('#156082')
      expect(one.textFrame.textRange.font.bold).toBe(true)
      expect(one.name).toMatch(/^Rectangle/)
    })
  })

  it('fails the sync with ItemNotFound for a missing item, and gives null objects for *OrNullObject', async () => {
    const deck = createPresentation([slide([shape({ id: 'a', tags: { bwb_element: 'callout' } })], { id: 's1' })])
    installFakePowerPoint(deck)
    await ppt(async (context) => {
      const shapes = context.presentation.slides.getItem('s1').shapes
      const tag = shapes.getItem('a').tags.getItemOrNullObject('BWB_ELEMENT')
      const missing = shapes.getItem('a').tags.getItemOrNullObject('BWB_NOTE')
      const none = shapes.getItemOrNullObject('zzz')
      tag.load('value')
      missing.load('value')
      none.load('id')
      await context.sync()
      expect(tag.isNullObject).toBe(false)
      expect(tag.value).toBe('callout')
      expect(missing.isNullObject).toBe(true)
      expect(() => missing.value).toThrow(/null object/)
      expect(none.isNullObject).toBe(true)
      shapes.getItem('zzz').load('id')
      await expect(context.sync()).rejects.toMatchObject({ code: 'ItemNotFound' })
    })
  })
})

describe('requirement sets', () => {
  it('answers isSetSupported from the model and fails newer APIs with ApiNotFound', async () => {
    const deck = createPresentation([slide([shape({ id: 'a' }), shape({ id: 'b' })])])
    select(deck, ['a', 'b'])
    installFakePowerPoint(deck, { apiVersion: '1.5' })
    expect(Office.context.requirements.isSetSupported('PowerPointApi', '1.5')).toBe(true)
    expect(Office.context.requirements.isSetSupported('PowerPointApi', '1.8')).toBe(false)
    expect(Office.context.requirements.isSetSupported('WordApi', '1.1')).toBe(false)
    await expect(ppt(async (context) => {
      context.presentation.getSelectedShapes().group()
      await context.sync()
    })).rejects.toMatchObject({ code: 'ApiNotFound' })
    await expect(ppt(async (context) => {
      context.presentation.getSelectedShapes().load('items/rotation')
      await context.sync()
    })).rejects.toMatchObject({ code: 'ApiNotFound' })
    expect(deck.slides[0]?.shapes).toHaveLength(2)
  })

  it('compares versions part by part, so 1.10 is newer than 1.9', () => {
    installFakePowerPoint(createPresentation(), { apiVersion: '1.10' })
    expect(Office.context.requirements.isSetSupported('PowerPointApi', '1.9')).toBe(true)
    expect(Office.context.requirements.isSetSupported('PowerPointApi', '1.11')).toBe(false)
  })
})

describe('selection', () => {
  it('lists selected shapes in z-order, and selected slides with the active slide first', async () => {
    const deck = createPresentation([slide([shape({ id: 'a' }), shape({ id: 'b' }), shape({ id: 'c' })], { id: 's1' }), slide([], { id: 's2' })])
    select(deck, ['c', 'a'])
    deck.selection.slideIds = ['s1', 's2']
    installFakePowerPoint(deck)
    await ppt(async (context) => {
      const shapes = context.presentation.getSelectedShapes()
      const slides = context.presentation.getSelectedSlides()
      shapes.load('items/id')
      slides.load('items/id')
      await context.sync()
      expect(shapes.items.map((s) => s.id)).toEqual(['a', 'c'])
      expect(slides.items.map((s) => s.id)).toEqual(['s1', 's2'])
    })
  })

  it('selects shapes and slides through the API', async () => {
    const deck = createPresentation([slide([shape({ id: 'a' })], { id: 's1' }), slide([shape({ id: 'b' })], { id: 's2' })])
    installFakePowerPoint(deck)
    await ppt(async (context) => {
      context.presentation.slides.getItem('s2').setSelectedShapes(['b'])
      await context.sync()
      expect(deck.selection).toMatchObject({ slideIds: ['s2'], shapeIds: ['b'] })
      context.presentation.setSelectedSlides(['s1'])
      await context.sync()
      expect(deck.selection).toMatchObject({ slideIds: ['s1'], shapeIds: [] })
      context.presentation.slides.getItem('s1').setSelectedShapes(['b'])
      await expect(context.sync()).rejects.toMatchObject({ code: 'InvalidArgument' })
    })
  })

  it('gives the selected text range, or a null object', async () => {
    const deck = createPresentation([slide([textBox('Hello world', { id: 't' })])])
    installFakePowerPoint(deck)
    await ppt(async (context) => {
      const none = context.presentation.getSelectedTextRangeOrNullObject()
      none.load('text')
      await context.sync()
      expect(none.isNullObject).toBe(true)
    })
    selectText(deck, 't', 6, 5)
    await ppt(async (context) => {
      const range = context.presentation.getSelectedTextRangeOrNullObject()
      range.load('text,start,length')
      await context.sync()
      expect([range.text, range.start, range.length]).toEqual(['world', 6, 5])
    })
  })

  it('fires DocumentSelectionChanged to the handlers that are added, and awaits them', async () => {
    const deck = createPresentation([slide([shape({ id: 'a' }), shape({ id: 'b' })])])
    installFakePowerPoint(deck)
    const seen: string[][] = []
    const handler = async (): Promise<void> => {
      await ppt(async (context) => {
        const shapes = context.presentation.getSelectedShapes()
        shapes.load('items/id')
        await context.sync()
        seen.push(shapes.items.map((s) => s.id))
      })
    }
    Office.context.document.addHandlerAsync(Office.EventType.DocumentSelectionChanged, handler)
    await clickShapes(deck, ['b'])
    await clickShapes(deck, ['a', 'b'])
    expect(seen).toEqual([['b'], ['a', 'b']])
    Office.context.document.removeHandlerAsync(Office.EventType.DocumentSelectionChanged, { handler })
    await fireSelectionChanged(deck)
    expect(seen).toHaveLength(2)
  })
})

describe('shapes', () => {
  it('adds shapes, text boxes, lines and tables, with names, ids and tags in upper case', async () => {
    const s1 = slide([], { id: 's1' })
    const deck = createPresentation([s1])
    installFakePowerPoint(deck)
    const id = await ppt(async (context) => {
      const shapes = context.presentation.slides.getItem('s1').shapes
      const box = shapes.addGeometricShape('Chevron', { left: 40, top: 108, width: 300, height: 64 })
      box.name = 'BwB process'
      box.tags.add('bwb_element', 'process')
      box.fill.setSolidColor('#022338')
      box.lineFormat.visible = false
      shapes.addTextBox('Source: x', { left: 40, top: 476, width: 700, height: 24 })
      shapes.addLine('Straight', { left: 40, top: 200, width: 880, height: 0 })
      shapes.addTable(2, 2, { values: [['A', 'B'], ['1', '2']], left: 40, top: 108 })
      box.load('id')
      await context.sync()
      return box.id
    })
    const [box, text, line, table] = s1.shapes
    expect(box).toMatchObject({ id, name: 'BwB process', type: 'GeometricShape', geometricShapeType: 'Chevron', left: 40, width: 300 })
    expect(box?.tags.get('BWB_ELEMENT')).toBe('process')
    expect(box?.fill).toMatchObject({ type: 'Solid', foregroundColor: '#022338' })
    expect(box?.line.visible).toBe(false)
    expect(text && textOf(text)).toBe('Source: x')
    expect(line).toMatchObject({ type: 'Line', connectorType: 'Straight', width: 880, textFrame: null })
    expect(table?.table?.values).toEqual([['A', 'B'], ['1', '2']])
  })

  it('rejects a table whose values do not match its size', async () => {
    installFakePowerPoint(createPresentation([slide([], { id: 's1' })]))
    await expect(ppt(async (context) => {
      context.presentation.slides.getItem('s1').shapes.addTable(2, 3, { values: [['a']] })
      await context.sync()
    })).rejects.toMatchObject({ code: 'InvalidArgument' })
  })

  it('changes the z-order and deletes shapes', async () => {
    const s1 = slide([shape({ id: 'a' }), shape({ id: 'b' }), shape({ id: 'c' })], { id: 's1' })
    const deck = createPresentation([s1])
    select(deck, ['a'])
    installFakePowerPoint(deck)
    await ppt(async (context) => {
      const shapes = context.presentation.slides.getItem('s1').shapes
      shapes.getItem('a').setZOrder('BringToFront')
      await context.sync()
      expect(s1.shapes.map((s) => s.id)).toEqual(['b', 'c', 'a'])
      shapes.getItem('a').setZOrder('SendBackward')
      const c = shapes.getItem('c')
      c.load('zOrderPosition')
      await context.sync()
      expect(c.zOrderPosition).toBe(2)
      shapes.getItem('a').delete()
      await context.sync()
    })
    expect(s1.shapes.map((s) => s.id)).toEqual(['b', 'c'])
    expect(deck.selection.shapeIds).toEqual([])
  })

  it('groups and ungroups, and moving a group moves its members', async () => {
    const s1 = slide([shape({ id: 'a', left: 0, top: 0 }), shape({ id: 'b', left: 200, top: 100 }), shape({ id: 'c' })], { id: 's1' })
    const deck = createPresentation([s1])
    select(deck, ['a', 'b'])
    installFakePowerPoint(deck)
    const groupId = await ppt(async (context) => {
      const g = context.presentation.getSelectedShapes().group()
      g.load('id,type,left,width')
      await context.sync()
      expect([g.type, g.left, g.width]).toEqual(['Group', 0, 300])
      g.left = 10
      await context.sync()
      return g.id
    })
    expect(s1.shapes.map((s) => s.id)).toEqual([groupId, 'c'])
    expect(deck.selection.shapeIds).toEqual([groupId])
    const g = s1.shapes[0]
    expect(g?.members?.map((m) => [m.id, m.left])).toEqual([['a', 10], ['b', 210]])
    await ppt(async (context) => {
      const shape0 = context.presentation.slides.getItem('s1').shapes.getItem(groupId)
      const members = shape0.group.shapes
      members.load('items/id')
      await context.sync()
      expect(members.items.map((m) => m.id)).toEqual(['a', 'b'])
      shape0.group.ungroup()
      await context.sync()
    })
    expect(s1.shapes.map((s) => s.id)).toEqual(['a', 'b', 'c'])
    expect(deck.selection.shapeIds).toEqual(['a', 'b'])
  })

  it('scales group members when the group is resized', () => {
    const g = group([shape({ left: 0, top: 0, width: 50, height: 50 }), shape({ left: 50, top: 0, width: 50, height: 50 })])
    expect(boundsOf(g)).toEqual({ left: 0, top: 0, width: 100, height: 50 })
  })

  it('fails text frames on shapes without text, and gives a null object from getTextFrameOrNullObject', async () => {
    const deck = createPresentation([slide([picture({ id: 'p' })], { id: 's1' })])
    installFakePowerPoint(deck)
    await ppt(async (context) => {
      const p = context.presentation.slides.getItem('s1').shapes.getItem('p')
      const frame = p.getTextFrameOrNullObject()
      frame.load('wordWrap')
      await context.sync()
      expect(frame.isNullObject).toBe(true)
      p.textFrame.load('wordWrap')
      await expect(context.sync()).rejects.toMatchObject({ code: 'InvalidArgument' })
    })
  })

  it('reads placeholder types, and fails placeholderFormat on other shapes', async () => {
    const deck = createPresentation([slide([placeholder('Title', 'A title', { id: 't' }), shape({ id: 'r' })], { id: 's1' })])
    installFakePowerPoint(deck)
    await ppt(async (context) => {
      const shapes = context.presentation.slides.getItem('s1').shapes
      const t = shapes.getItem('t')
      t.load('placeholderFormat/type')
      await context.sync()
      expect(t.placeholderFormat.type).toBe('Title')
      shapes.getItem('r').placeholderFormat.load('type')
      await expect(context.sync()).rejects.toMatchObject({ code: 'InvalidArgument' })
    })
  })
})
