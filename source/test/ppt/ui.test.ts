import { createElement } from 'react'
import { renderToStaticMarkup } from 'react-dom/server'
import { afterEach, describe, expect, it } from 'vitest'
import { ArrangeTab } from '../../src/ui/ppt/ArrangeTab.tsx'
import { ElementsPanel } from '../../src/ui/ppt/ElementsPanel.tsx'
import { createPresentation, installFakePowerPoint, uninstallFakePowerPoint } from '../fakePpt/index.ts'

afterEach(uninstallFakePowerPoint)

const render = (component: () => unknown): string => renderToStaticMarkup(createElement(component as () => null))

describe('the arrange tab', () => {
  it('renders every tool with the reference choice and the anchor', () => {
    const html = render(ArrangeTab)
    for (const text of ['Nothing selected', 'Set anchor', 'Line up with', 'Selection', 'Anchor', 'Margins', 'Content', 'Spread out', 'Even out gaps',
      'Same width', 'Same height', 'Same size', 'Read from selection', 'Apply', 'Side by side', 'Top to bottom', 'Tidy up', 'Snap', 'Swap',
      'Left edge', 'Bottom edge', 'Nudge left', 'To front', 'Group', 'Ungroup', 'Autofit off', 'Margins 0.2 cm']) {
      expect(html, text).toContain(text)
    }
    expect(html).not.toMatch(/[\u2014!]/)
  })

  it('disables what this PowerPoint cannot do, and says why', () => {
    installFakePowerPoint(createPresentation(), { apiVersion: '1.7' })
    const html = render(ArrangeTab)
    expect(html).toContain('Changing the order of shapes or grouping them needs a newer PowerPoint.')
    expect(html).toContain('This PowerPoint does not report rotation')
    installFakePowerPoint(createPresentation(), { apiVersion: '1.4' })
    expect(render(ArrangeTab)).toContain('Arranging shapes needs a newer PowerPoint.')
  })
})

describe('the elements panel', () => {
  it('renders the elements, marks and review notes, leaving stickers to the Slides tab', () => {
    const html = render(ElementsPanel)
    for (const text of ['Callout', 'Key figures', 'Quote', 'Process, three steps', 'Two-by-two matrix', 'Timeline', 'Numbered markers', 'Source line',
      'BwB table', 'Harvey balls', 'On track', 'Amber dot, for trackers only', 'Add a note', 'Remove all notes']) {
      expect(html, text).toContain(text)
    }
    expect(html).not.toContain('Not for distribution')
    expect(html).not.toMatch(/[\u2014!]/)
  })

  it('marks the table as needing a newer PowerPoint where it cannot be made', () => {
    installFakePowerPoint(createPresentation(), { apiVersion: '1.7' })
    expect(render(ElementsPanel)).toContain('This element needs a newer PowerPoint.')
  })
})
