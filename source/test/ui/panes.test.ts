import { createElement } from 'react'
import { renderToStaticMarkup } from 'react-dom/server'
import { afterEach, describe, expect, it } from 'vitest'
import { BrandTab } from '../../src/ui/tabs/BrandTab.tsx'
import { SlideInsertTab } from '../../src/ui/tabs/SlideInsertTab.tsx'
import { SlidesPanel } from '../../src/ui/tabs/SlidesPanel.tsx'
import { tabsFor } from '../../src/ui/tabSets.tsx'
import { ToastProvider } from '../../src/ui/toast.tsx'
import { createPresentation, installFakePowerPoint, uninstallFakePowerPoint } from '../fakePpt/index.ts'

afterEach(uninstallFakePowerPoint)

const render = (component: () => unknown): string =>
  renderToStaticMarkup(createElement(ToastProvider, null, createElement(component as () => null)))

describe('the pane in each app', () => {
  it('opens Word and PowerPoint on their own tabs', () => {
    expect(tabsFor('word').map((t) => t.id)).toEqual(['start', 'write', 'insert', 'tools', 'check', 'brand'])
    expect(tabsFor('powerpoint').map((t) => t.id)).toEqual(['slides', 'arrange', 'insert', 'check', 'brand'])
  })

  it('starts the Slides tab with a new deck, and puts the stickers there', () => {
    installFakePowerPoint(createPresentation())
    const html = render(SlidesPanel)
    for (const text of ['Start a new BwB deck', 'Draft', 'For discussion', 'Not for distribution', 'Remove every sticker in the deck', 'Slide library']) {
      expect(html, text).toContain(text)
    }
    expect(html.indexOf('Start a new BwB deck')).toBeLessThan(html.indexOf('Slide library'))
  })

  it('gathers elements, charts, icons, photos and logos on the PowerPoint Insert tab', () => {
    installFakePowerPoint(createPresentation())
    const html = render(SlideInsertTab)
    for (const id of ['section-charts', 'section-icons', 'section-photos']) expect(html, id).toContain(`id="${id}"`)
    expect(html).toContain('Logo in the footer')
  })

  it('shows slide rules and shape colour buttons on the PowerPoint Brand tab', () => {
    installFakePowerPoint(createPresentation())
    const html = render(BrandTab)
    for (const text of ['Type on slides', 'The slide grid', 'Slide furniture', 'Fill the selected shapes navy', 'Colour the selected text charcoal']) {
      expect(html, text).toContain(text)
    }
    expect(html).not.toContain('Document furniture')
  })

  it('keeps the document rules on the Word Brand tab', () => {
    const html = render(BrandTab)
    expect(html).toContain('Document furniture')
    expect(html).not.toContain('The slide grid')
  })

  it('never uses an em dash or an exclamation mark in the copy', () => {
    installFakePowerPoint(createPresentation())
    for (const component of [SlidesPanel, SlideInsertTab, BrandTab]) expect(render(component)).not.toMatch(/[\u2014!]/)
  })
})
