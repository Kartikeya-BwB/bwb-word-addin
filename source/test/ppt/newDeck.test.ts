import { afterEach, describe, expect, it, vi } from 'vitest'
import { createBwbDeck } from '../../src/ppt/newDeck.ts'
import { createPresentation, installFakePowerPoint, uninstallFakePowerPoint } from '../fakePpt/index.ts'

afterEach(() => {
  uninstallFakePowerPoint()
  vi.unstubAllGlobals()
})

describe('a new BwB deck', () => {
  it('opens the starter deck as a new presentation', async () => {
    const model = createPresentation()
    installFakePowerPoint(model)
    const fetch = vi.fn(async () => new Response(new Uint8Array([80, 75, 3, 4])))
    vi.stubGlobal('fetch', fetch)
    expect(await createBwbDeck()).toMatch(/new BwB deck is open/)
    expect(fetch).toHaveBeenCalledWith('./templates/bwb-deck.pptx')
    expect(model.createdPresentations).toEqual([btoa('PK\x03\x04')])
  })

  it('says so when the template cannot be loaded', async () => {
    installFakePowerPoint(createPresentation())
    vi.stubGlobal('fetch', vi.fn(async () => new Response('', { status: 404 })))
    await expect(createBwbDeck()).rejects.toThrow('The BwB template could not be loaded.')
  })

  it('needs PowerPoint', async () => {
    await expect(createBwbDeck()).rejects.toThrow('Open this in PowerPoint')
  })
})
