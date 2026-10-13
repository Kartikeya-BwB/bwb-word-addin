import { afterEach, describe, expect, it } from 'vitest'
import { insertPictureHere, insertSvgHere } from '../../src/media/insert.ts'
import { createPresentation, findShape, installFakePowerPoint, slide, uninstallFakePowerPoint } from '../fakePpt/index.ts'

afterEach(uninstallFakePowerPoint)

const deck = () => createPresentation([slide([], { id: 's1' })])

describe('placing pictures on a slide', () => {
  it('centres a photo and gives it alt text, so the deck check does not flag it', async () => {
    const model = deck()
    installFakePowerPoint(model)
    await insertPictureHere({ bytes: new Uint8Array([137, 80, 78, 71]), widthPt: 200, heightPt: 100, altText: 'Solar panels in Kenya' })
    const [insert] = model.dataInserts
    expect(insert?.coercionType).toBe('image')
    expect(insert?.options).toMatchObject({ imageWidth: 200, imageHeight: 100, imageLeft: 380, imageTop: 220 })
    const shape = findShape(model, insert?.shapeId ?? '')
    expect(shape?.altTextDescription).toBe('Solar panels in Kenya')
  })

  it('tags a chart so the add-in can find it again, at the place asked for', async () => {
    const model = deck()
    installFakePowerPoint(model)
    await insertSvgHere({ svg: '<svg xmlns="http://www.w3.org/2000/svg"/>', widthPt: 880, heightPt: 360, altText: 'Column chart', left: 40, top: 108, tags: { BWB_CHART: 'Funding' } })
    const [insert] = model.dataInserts
    expect(insert?.coercionType).toBe('xmlSvg')
    expect(insert?.options).toMatchObject({ imageLeft: 40, imageTop: 108 })
    const shape = findShape(model, insert?.shapeId ?? '')
    expect(shape?.tags.get('BWB_CHART')).toBe('Funding')
    expect(shape?.altTextDescription).toBe('Column chart')
  })

  it('still places the picture on an older PowerPoint that cannot set alt text', async () => {
    const model = deck()
    installFakePowerPoint(model, { apiVersion: '1.5' })
    await insertPictureHere({ bytes: new Uint8Array([1]), widthPt: 50, heightPt: 50, altText: 'Logo', tags: { BWB_ELEMENT: 'logo' } })
    const shape = findShape(model, model.dataInserts[0]?.shapeId ?? '')
    expect(shape?.altTextDescription).toBe('')
    expect(shape?.tags.get('BWB_ELEMENT')).toBe('logo')
  })

  it('reports a failed insert in plain words', async () => {
    const model = deck()
    model.failNextDataInsert = 'The slide is locked'
    installFakePowerPoint(model)
    await expect(insertPictureHere({ bytes: new Uint8Array([1]), widthPt: 50, heightPt: 50, altText: 'Logo' }))
      .rejects.toThrow('PowerPoint did not insert the picture (The slide is locked).')
  })
})
