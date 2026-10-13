import { describe, expect, it } from 'vitest'
import { contrastNeeded, contrastRatio } from '../../../src/ppt/check/colour.ts'
import { deckBulkFixes, runDeckCheck, shownShapes } from '../../../src/ppt/check/engine.ts'
import { addressesIn, internalLink } from '../../../src/ppt/check/links.ts'
import { accessFindings, colourBehind, hiddenFindings } from '../../../src/ppt/check/rulesAccess.ts'
import type { DeckFinding, DeckSnapshot } from '../../../src/ppt/check/types.ts'
import { deck, font, picture, shape, slide, text, title } from './build.ts'

const rules = (findings: DeckFinding[]): string[] => findings.map((f) => f.rule)
const white = (shapes: Parameters<typeof slide>[0], over: Parameters<typeof slide>[1] = {}): DeckSnapshot =>
  deck([slide(shapes, { background: 'FFFFFF', ...over })])

describe('contrast', () => {
  it('follows WCAG: 21:1 for black on white, 4.5:1 for text, 3:1 for large text', () => {
    expect(contrastRatio('000000', 'FFFFFF')).toBeCloseTo(21, 5)
    expect(contrastRatio('FFFFFF', '022338')).toBeGreaterThan(14)
    expect(contrastRatio('328840', 'FFFFFF')).toBeLessThan(4.5)
    expect([contrastNeeded(14, false), contrastNeeded(18, false), contrastNeeded(14, true), contrastNeeded(null, null)]).toEqual([4.5, 3, 3, 4.5])
  })

  it('flags text too pale for its background, and offers the clearest text colour', () => {
    const found = accessFindings(white([text('Pale grey', { font: font({ colour: '9A9A9A' }) })]))
    expect(found.map((f) => [f.rule, f.title])).toEqual([['brand/contrast', 'Low contrast (2.8:1)']])
    expect(found[0]?.detail).toContain('#9A9A9A text on white needs a contrast of at least 4.5:1')
    expect(found[0]?.fix).toMatchObject({ kind: 'text-colour', colour: '000000', spans: [{ start: 0, length: 9 }] })
    expect(found[0]?.fixLabel).toBe('Change to Black')
  })

  it('lets green through only when it is large, as the BwB colour rules do', () => {
    const green = (size: number, bold = false) => text(`Green at ${size}`, { font: font({ colour: '328840', size, bold }) })
    expect(accessFindings(white([green(14)])).map((f) => f.excerpt)).toEqual(['Green at 14'])
    expect(accessFindings(white([green(18), green(14, true)]))).toEqual([])
  })

  it('reads the colour behind: the shape’s fill, then a filled shape under it, then the slide', () => {
    const panel = shape({ id: 'panel', type: 'GeometricShape', text: '', left: 0, top: 0, width: 500, height: 300, fill: { type: 'Solid', colour: '022338' } })
    const onPanel = text('White on navy', { id: 'on', left: 40, top: 40, width: 200, height: 40, font: font({ colour: 'FFFFFF' }) })
    const offPanel = text('White on white', { id: 'off', left: 600, top: 40, width: 200, height: 40, font: font({ colour: 'FFFFFF' }) })
    const snap = white([panel, onPanel, offPanel])
    const s = snap.slides[0]
    if (!s) throw new Error('no slide')
    expect([colourBehind(s, 1), colourBehind(s, 2)]).toEqual(['022338', 'FFFFFF'])
    expect(accessFindings(snap).map((f) => f.excerpt)).toEqual(['White on white'])
  })

  it('stays quiet when it cannot know the background', () => {
    const pale = (over = {}) => text('Pale', { left: 40, top: 40, width: 100, height: 20, font: font({ colour: 'DDDDDD' }), ...over })
    const photo = picture({ left: 0, top: 0, width: 960, height: 540 })
    expect(accessFindings(deck([slide([pale()])]))).toEqual([])
    expect(accessFindings(white([photo, pale()]))).toEqual([])
    expect(accessFindings(white([pale({ fill: { type: 'Solid', colour: 'FFFFFF', transparency: 0.5 } })]))).toEqual([])
    expect(accessFindings(white([pale({ tags: { BWB_RAG: 'amber' } })]))).toEqual([])
  })

  it('leaves lime on a light fill to the lime rule', () => {
    const lime = text('Lime', { font: font({ colour: 'BED62F' }), fill: { type: 'Solid', colour: 'FFFFFF' } })
    expect(rules(accessFindings(white([lime])))).toEqual([])
    expect(rules(runDeckCheck(white([title('A title'), lime])).findings)).toContain('brand/lime-on-light')
  })
})

describe('hidden shapes', () => {
  it('are flagged once, a hidden group as a whole, with a fix that deletes them', () => {
    const snap = white([
      title('A title'),
      text('Old version of the chart', { id: 'old', visible: false }),
      shape({ id: 'g', type: 'Group', text: null, visible: false }),
      text('Inside the hidden group', { id: 'in', level: 1, path: ['g', 'in'], visible: true }),
    ])
    const found = hiddenFindings(snap)
    expect(found.map((f) => f.location?.kind === 'shape' && f.location.path.join('/'))).toEqual(['old', 'g'])
    expect(found[0]).toMatchObject({ group: 'send', fix: { kind: 'delete-shape' }, excerpt: 'Old version of the chart' })
  })

  it('are left out of every other rule', () => {
    const snap = white([title('A title'), text('Hidden Arial at 6 pt', { visible: false, font: font({ name: 'Arial', size: 6 }) })])
    expect(shownShapes(snap).slides[0]?.shapes).toHaveLength(1)
    expect(rules(runDeckCheck(snap).findings)).toEqual(['send/hidden-shape'])
  })
})

describe('internal links', () => {
  it('know files, network folders, intranet addresses and cloud storage from public pages', () => {
    expect(internalLink('file:///C:/Users/ana/report.docx')).toBe('file')
    expect(internalLink('C:\\Users\\ana\\report.docx')).toBe('file')
    expect(internalLink('\\\\bwb-files\\projects\\data.xlsx')).toBe('network')
    expect(internalLink('http://intranet/policies')).toBe('intranet')
    expect(internalLink('https://wiki.corp/page')).toBe('intranet')
    expect(internalLink('http://192.168.1.20/report')).toBe('intranet')
    expect(internalLink('https://bwb.sharepoint.com/sites/qa/doc.pptx')).toBe('storage')
    expect(internalLink('https://bwb-my.sharepoint.com/personal/ana/doc.pptx')).toBe('storage')
    expect(internalLink('https://1drv.ms/x/abc')).toBe('storage')
    expect(internalLink('https://www.bwb.earth/insights')).toBeNull()
    expect(internalLink('mailto:qa@bwb.earth')).toBeNull()
    expect(internalLink('not a link')).toBeNull()
  })

  it('finds addresses typed in text, without trailing punctuation', () => {
    expect(addressesIn('See https://bwb.sharepoint.com/x/y. Or \\\\server\\share\\f.xlsx, and C:\\temp\\a.docx; done'))
      .toEqual(['https://bwb.sharepoint.com/x/y', '\\\\server\\share\\f.xlsx', 'C:\\temp\\a.docx'])
  })

  it('are flagged from the slide’s hyperlinks and from its text, once per address', () => {
    const box = text('Data: https://bwb.sharepoint.com/sites/x/data.xlsx', { id: 'box' })
    const snap = white([title('A title'), box, shape({ id: 'btn', type: 'GeometricShape', text: '' })], {
      links: [
        { address: 'https://bwb.sharepoint.com/sites/x/data.xlsx', shapeId: null },
        { address: 'http://intranet/team', shapeId: 'btn' },
        { address: 'https://www.bwb.earth', shapeId: null },
      ],
    })
    const found = accessFindings(snap).filter((f) => f.rule === 'send/internal-link')
    expect(found.map((f) => [f.excerpt, f.location?.kind === 'shape' ? f.location.path[0] : 'slide'])).toEqual([
      ['https://bwb.sharepoint.com/sites/x/data.xlsx', 'slide'], ['http://intranet/team', 'btn'],
    ])
    expect(found[1]?.detail).toContain('internal web address')
  })
})

describe('fixing in bulk', () => {
  it('counts deleting an empty slide as safe', () => {
    const snap = deck([slide([title('Kept')]), slide([title('', { id: 'empty-title' })])])
    expect(deckBulkFixes(runDeckCheck(snap).findings).map((f) => f.fix?.kind)).toContain('delete-slide')
  })
})
