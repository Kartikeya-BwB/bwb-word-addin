import { createElement } from 'react'
import { renderToStaticMarkup } from 'react-dom/server'
import { afterEach, describe, expect, it } from 'vitest'
import { DeckCheckTab } from '../../../src/ui/ppt/DeckCheckTab.tsx'
import { SlidesTab } from '../../../src/ui/ppt/SlidesTab.tsx'
import { createPresentation, installFakePowerPoint, uninstallFakePowerPoint } from '../../fakePpt/index.ts'

afterEach(uninstallFakePowerPoint)

const slidesTab = (stickers?: string): string =>
  renderToStaticMarkup(createElement(SlidesTab, stickers ? { stickers: createElement('p', null, stickers) } : {}))
const deckCheckTab = (): string => renderToStaticMarkup(createElement(DeckCheckTab))

describe('the Slides tab', () => {
  it('renders the library, storyline, sections, stickers and copy sections, with the anchors the ribbon opens', () => {
    const html = slidesTab()
    for (const text of [
      'Slide library', 'Loading the slide library', 'Storyline', 'Read the storyline', 'Sections and agenda', 'Start a section at this slide',
      'Insert agenda slide', 'Section trackers', 'Stickers', 'Use Sticker on the BwB ribbon tab', 'Copy slides to a new deck',
    ]) expect(html, text).toContain(text)
    for (const id of ['library', 'storyline', 'sections', 'stickers', 'copy']) expect(html, id).toContain(`id="section-${id}"`)
    expect(html).not.toMatch(/[\u2014!]/)
  })

  it('shows the sticker buttons it is given', () => {
    expect(slidesTab('Sticker buttons here')).toContain('Sticker buttons here')
  })

  it('says plainly what this PowerPoint is too old for', () => {
    installFakePowerPoint(createPresentation(), { apiVersion: '1.9' })
    expect(slidesTab()).toContain('Copying slides to a new presentation needs a newer PowerPoint.')
    installFakePowerPoint(createPresentation(), { apiVersion: '1.4' })
    expect(slidesTab()).toContain('Starting a section at the current slide needs a newer PowerPoint.')
    installFakePowerPoint(createPresentation(), { apiVersion: '1.2' })
    const old = slidesTab()
    expect(old).toContain('The Slides tab needs a newer PowerPoint.')
    expect(old).toContain('Inserting library slides needs a newer PowerPoint.')
  })
})

describe('the Pre-QA check tab', () => {
  it('renders the check and the QA note, with nothing run yet', () => {
    const html = deckCheckTab()
    for (const text of ['Pre-QA check', 'Run the check', 'Send to QA', 'Copy QA note']) expect(html, text).toContain(text)
    expect(html).not.toMatch(/[\u2014!]/)
  })

  it('says plainly when this PowerPoint is too old to check the deck', () => {
    installFakePowerPoint(createPresentation(), { apiVersion: '1.3' })
    expect(deckCheckTab()).toContain('The Pre-QA check needs a newer PowerPoint.')
  })
})
