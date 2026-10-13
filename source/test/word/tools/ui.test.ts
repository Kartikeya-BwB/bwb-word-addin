import { createElement } from 'react'
import { renderToStaticMarkup } from 'react-dom/server'
import { afterEach, describe, expect, it } from 'vitest'
import { ToolsTab } from '../../../src/ui/tabs/ToolsTab.tsx'
import type { FakeModel } from '../../fakeWord/model.ts'
import { createModel, para } from '../../fakeWord/model.ts'
import { installFakeWord, uninstallFakeWord } from '../../fakeWord/install.ts'

afterEach(uninstallFakeWord)

const render = (over: Partial<FakeModel> = {}): string => {
  installFakeWord(createModel([para('Body')], over))
  return renderToStaticMarkup(createElement(ToolsTab))
}

describe('ToolsTab', () => {
  it('has the five sections, with the navigator where the ribbon opens it', () => {
    const html = render({ platform: 'PC', desktopApiVersion: 1.4 })
    for (const id of ['pictures', 'tables', 'text', 'navigator', 'review']) expect(html).toContain(`id="section-${id}"`)
    for (const title of ['Pictures', 'Tables', 'Text', 'Navigator', 'Review']) expect(html).toContain(`<h2>${title}</h2>`)
    expect(html).toContain('Fit to text width (16.6 cm)')
    expect(html).toContain('put the cursor where it goes, then choose Ref')
    expect(html).not.toContain('Needs Word for Windows or Mac.')
    expect(html).not.toMatch(/[\u2014!]/)
  })

  it('disables the desktop-only tools on the web and says why', () => {
    const html = render({ platform: 'OfficeOnline', desktopApiVersion: 0 })
    expect(html).toContain('Needs Word for Windows or Mac.')
    expect(html).toContain('Cross-references need Word for Windows or Mac: Word on the web cannot add fields.')
    expect(html).not.toContain('then choose Ref')
  })

  it('disables the review tools in a Word without WordApi 1.4', () => {
    const html = render({ apiVersion: 1.3 })
    expect(html).toContain('Needs a newer Word (WordApi 1.4): update Microsoft 365.')
    expect(html).toContain('Needs a newer Word (WordApi 1.6): update Microsoft 365.')
  })
})
