import { afterEach, describe, expect, it } from 'vitest'
import {
  createPresentation, installFakePowerPoint, picture, placeholder, run, shape, slide, tableShape, textBox, textFrame, textOf,
  uninstallFakePowerPoint,
} from './index.ts'

afterEach(uninstallFakePowerPoint)

const ppt = <T>(fn: (context: PowerPoint.RequestContext) => Promise<T>): Promise<T> => PowerPoint.run(fn)

describe('text', () => {
  it('writes a substring in the font of the text it replaces, keeping other formatting', async () => {
    const t = textBox('', { id: 't', textFrame: textFrame('', { runs: [run('Plain '), run('bold', { bold: true }), run(' end')] }) })
    const deck = createPresentation([slide([t], { id: 's1' })])
    installFakePowerPoint(deck)
    await ppt(async (context) => {
      const range = context.presentation.slides.getItem('s1').shapes.getItem('t').textFrame.textRange
      range.getSubstring(6, 4).text = 'heavy'
      range.getSubstring(0, 5).text = 'Clear'
      const whole = range.font
      whole.load('bold')
      await context.sync()
      expect(whole.bold).toBeNull()
    })
    expect(textOf(t)).toBe('Clear heavy end')
    expect(t.textFrame?.runs.map((r) => [r.text, r.font.bold])).toEqual([['Clear ', false], ['heavy', true], [' end', false]])
  })

  it('sets fonts, alignment and frame settings', async () => {
    const t = shape({ id: 't', text: 'Key figure\nLabel' })
    const deck = createPresentation([slide([t], { id: 's1' })])
    installFakePowerPoint(deck)
    await ppt(async (context) => {
      const frame = context.presentation.slides.getItem('s1').shapes.getItem('t').textFrame
      frame.textRange.font.name = 'Aptos'
      frame.textRange.getSubstring(0, 10).font.color = '#bed62f'
      frame.textRange.paragraphFormat.horizontalAlignment = 'Center'
      frame.verticalAlignment = 'Middle'
      frame.autoSizeSetting = 'AutoSizeNone'
      frame.leftMargin = 0
      await context.sync()
    })
    expect(t.textFrame?.runs.map((r) => [r.text, r.font.color, r.font.name])).toEqual([['Key figure', '#BED62F', 'Aptos'], ['\nLabel', '#FFFFFF', 'Aptos']])
    expect(t.textFrame).toMatchObject({ horizontalAlignment: 'Center', verticalAlignment: 'Middle', leftMargin: 0 })
  })

  it('rejects a substring outside the text', async () => {
    installFakePowerPoint(createPresentation([slide([textBox('Short', { id: 't' })], { id: 's1' })]))
    await expect(ppt(async (context) => {
      context.presentation.slides.getItem('s1').shapes.getItem('t').textFrame.textRange.getSubstring(3, 10).text = 'x'
      await context.sync()
    })).rejects.toMatchObject({ code: 'InvalidArgument' })
  })
})

describe('presentation', () => {
  it('adds slides from a layout, copying its placeholders, and moves and deletes slides', async () => {
    const deck = createPresentation([slide([], { id: 's1' })])
    installFakePowerPoint(deck)
    await ppt(async (context) => {
      context.presentation.slides.add({ layoutId: 'layout-title-only' })
      await context.sync()
      const slides = context.presentation.slides
      slides.load('items/id')
      await context.sync()
      const added = slides.items[1]
      added?.moveTo(0)
      await context.sync()
      context.presentation.slides.getItem('s1').delete()
    })
    expect(deck.slides).toHaveLength(1)
    expect(deck.slides[0]?.layoutId).toBe('layout-title-only')
    expect(deck.slides[0]?.shapes.map((s) => s.placeholderType)).toEqual(['Title'])
  })

  it('records insertSlidesFromBase64 and inserts the slides a test provides', async () => {
    const deck = createPresentation([slide([], { id: 's1' }), slide([], { id: 's2' })], { slidesFromBase64: () => [slide([], { id: 'new' })] })
    installFakePowerPoint(deck)
    await ppt(async (context) => {
      context.presentation.insertSlidesFromBase64('AAAA', { targetSlideId: 's1', formatting: 'UseDestinationTheme' })
      await context.sync()
      context.presentation.insertSlidesFromBase64('BBBB', { targetSlideId: 'nope' })
      await expect(context.sync()).rejects.toMatchObject({ code: 'SlideNotFound' })
    })
    expect(deck.slideInserts).toEqual([{ base64: 'AAAA', options: { targetSlideId: 's1', formatting: 'UseDestinationTheme' } }])
    expect(deck.slides.map((s) => s.id)).toEqual(['s1', 'new', 's2'])
  })

  it('reads and writes the page setup and presentation tags', async () => {
    const deck = createPresentation()
    installFakePowerPoint(deck)
    await ppt(async (context) => {
      const page = context.presentation.pageSetup
      page.load('slideWidth,slideHeight')
      context.presentation.tags.add('bwb_layout', 'v1')
      await context.sync()
      expect([page.slideWidth, page.slideHeight]).toEqual([960, 540])
    })
    expect(deck.tags.get('BWB_LAYOUT')).toBe('v1')
  })
})

describe('Office.context.document', () => {
  it('inserts pictures and SVG on the active slide, selected, and records the call', async () => {
    const s1 = slide([], { id: 's1' })
    const deck = createPresentation([s1])
    installFakePowerPoint(deck)
    const status = await new Promise((resolve) => {
      Office.context.document.setSelectedDataAsync('<svg/>', {
        coercionType: Office.CoercionType.XmlSvg, imageLeft: 10, imageTop: 20, imageWidth: 30, imageHeight: 30,
      }, (r) => resolve(r.status))
    })
    expect(status).toBe(Office.AsyncResultStatus.Succeeded)
    expect(s1.shapes[0]).toMatchObject({ type: 'Graphic', left: 10, top: 20, width: 30 })
    expect(deck.selection.shapeIds).toEqual([s1.shapes[0]?.id])
    expect(deck.dataInserts[0]).toMatchObject({ data: '<svg/>', coercionType: 'xmlSvg', shapeId: s1.shapes[0]?.id })
    deck.failNextDataInsert = 'No room'
    const failed = await new Promise((resolve) => {
      Office.context.document.setSelectedDataAsync('AAAA', { coercionType: Office.CoercionType.Image }, (r) => resolve(r.error?.message))
    })
    expect(failed).toBe('No room')
    expect(s1.shapes).toHaveLength(1)
  })

  it('reports PowerPoint as the host', () => {
    installFakePowerPoint(createPresentation())
    expect(Office.context.host).toBe(Office.HostType.PowerPoint)
    expect(PowerPoint.ShapeZOrder.bringToFront).toBe('BringToFront')
  })

  it('reads tables and records cell formatting', async () => {
    const t = tableShape([['A', 'B']], { id: 't' })
    installFakePowerPoint(createPresentation([slide([t], { id: 's1' })]))
    await ppt(async (context) => {
      const table = context.presentation.slides.getItem('s1').shapes.getItem('t').getTable()
      table.load('values,rowCount,columnCount')
      const cell = table.getCellOrNullObject(0, 1)
      cell.font.bold = true
      cell.fill.setSolidColor('#022338')
      await context.sync()
      expect([table.values, table.rowCount, table.columnCount]).toEqual([[['A', 'B']], 1, 2])
    })
    expect(t.table?.cells['0,1']).toEqual({ font: { bold: true }, fill: { color: '#022338' } })
  })
})

describe('for the slide tools and the deck check', () => {
  it('reads and writes isDecorative (PowerPointApi 1.10)', async () => {
    const pic = picture({ id: 'p', isDecorative: true })
    const deck = createPresentation([slide([pic, shape({ id: 'r' })], { id: 's1' })])
    installFakePowerPoint(deck)
    await ppt(async (context) => {
      const shapes = context.presentation.slides.getItem('s1').shapes
      shapes.load('items/id,items/isDecorative')
      await context.sync()
      expect(shapes.items.map((s) => s.isDecorative)).toEqual([true, false])
      shapes.getItem('r').isDecorative = true
      await context.sync()
    })
    expect(deck.slides[0]?.shapes.map((s) => s.isDecorative)).toEqual([true, true])
    installFakePowerPoint(deck, { apiVersion: '1.9' })
    await expect(ppt(async (context) => {
      context.presentation.slides.getItem('s1').shapes.getItem('p').load('isDecorative')
      await context.sync()
    })).rejects.toMatchObject({ code: 'ApiNotFound' })
  })

  it('records bullets shown or hidden per text frame', async () => {
    const body = placeholder('Body', 'Point one\nPoint two', { id: 'b' })
    const box = textBox('Plain', { id: 't' })
    installFakePowerPoint(createPresentation([slide([body, box], { id: 's1' })]))
    await ppt(async (context) => {
      const shapes = context.presentation.slides.getItem('s1').shapes
      const bullets = shapes.getItem('b').textFrame.textRange.paragraphFormat.bulletFormat
      bullets.load('visible')
      await context.sync()
      expect(bullets.visible).toBe(true)
      bullets.visible = false
      shapes.getItem('t').textFrame.textRange.paragraphFormat.bulletFormat.visible = true
      await context.sync()
    })
    expect([body.textFrame?.bulletVisible, box.textFrame?.bulletVisible]).toEqual([false, true])
  })

  it('exports slides as a presentation and records which (PowerPointApi 1.10)', async () => {
    const deck = createPresentation([slide([], { id: 's1' }), slide([], { id: 's2' }), slide([], { id: 's3' })])
    deck.selection.slideIds = ['s3']
    installFakePowerPoint(deck)
    await ppt(async (context) => {
      const all = context.presentation.slides.exportAsBase64Presentation(['s1', context.presentation.slides.getItem('s2')])
      const selected = context.presentation.getSelectedSlides().exportAsBase64Presentation()
      await context.sync()
      expect(atob(all.value)).toBe('fake-presentation:s1,s2')
      expect(selected.value.length).toBeGreaterThan(0)
      context.presentation.slides.exportAsBase64Presentation(['nope'])
      await expect(context.sync()).rejects.toMatchObject({ code: 'ItemNotFound' })
    })
    expect(deck.slideExports).toEqual([['s1', 's2'], ['s3']])
  })

  it('records PowerPoint.createPresentation calls', async () => {
    const deck = createPresentation()
    installFakePowerPoint(deck)
    await PowerPoint.createPresentation('UEsDBA==')
    await PowerPoint.createPresentation()
    expect(deck.createdPresentations).toEqual(['UEsDBA==', null])
  })

  it('keeps the ids a test gives the slides that insertSlidesFromBase64 inserts', async () => {
    const deck = createPresentation([slide([], { id: 's1' })], {
      slidesFromBase64: () => [slide([shape({ id: 'title-1' })], { id: 'lib-1' }), slide([], { id: 'lib-2' })],
    })
    installFakePowerPoint(deck)
    const ids = await ppt(async (context) => {
      context.presentation.insertSlidesFromBase64('AAAA', { targetSlideId: 's1' })
      await context.sync()
      const slides = context.presentation.slides
      slides.load('items/id')
      const shapes = context.presentation.slides.getItem('lib-1').shapes
      shapes.load('items/id')
      await context.sync()
      return [slides.items.map((s) => s.id), shapes.items.map((s) => s.id)]
    })
    expect(ids).toEqual([['s1', 'lib-1', 'lib-2'], ['title-1']])
  })
})

describe('builders', () => {
  it('keep the text given to textBox and placeholder when a font is given too', () => {
    expect(textOf(textBox('Note', { font: { name: 'Arial' } }))).toBe('Note')
    const title = placeholder('Title', 'A title', { font: { name: 'Calibri' } })
    expect([textOf(title), title.textFrame?.runs[0]?.font.name, title.textFrame?.runs[0]?.font.size]).toEqual(['A title', 'Calibri', 24])
    expect(textOf(shape({ text: 'Box', font: { bold: true } }))).toBe('Box')
  })
})
