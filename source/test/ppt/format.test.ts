import { afterEach, describe, expect, it } from 'vitest'
import {
  STANDARD_MARGIN, applyBwbFonts, applyFill, applyOutline, applyTextColour, cleanSelectedText, setAutofitOff, setTextMargins, setVerticalAnchor,
  setWordWrap, wordingFixes,
} from '../../src/ppt/format.ts'
import type { FakePresentation, FakeShape } from '../fakePpt/index.ts'
import {
  createPresentation, group, installFakePowerPoint, line, picture, placeholder, run, select, shape, slide, tableShape, textBox, textFrame, textOf,
  uninstallFakePowerPoint,
} from '../fakePpt/index.ts'

afterEach(uninstallFakePowerPoint)

function deck(shapes: FakeShape[], selected: string[], apiVersion = '1.10'): FakePresentation {
  const d = createPresentation([slide(shapes, { id: 's1' })], { apiVersion })
  select(d, selected)
  installFakePowerPoint(d)
  return d
}

describe('fills, text colours and outlines', () => {
  it('fills shapes in navy with white text, and clears fills', async () => {
    const a = shape({ id: 'a', text: 'One' })
    const t = textBox('Two', { id: 't' })
    deck([a, t, picture({ id: 'p' })], ['a', 't', 'p'])
    expect(await applyFill('navy')).toBe('Navy fill, white text, on 2 shapes.')
    expect([a.fill, t.fill].map((f) => [f.type, f.foregroundColor])).toEqual([['Solid', '#022338'], ['Solid', '#022338']])
    expect([a, t].map((s) => s.textFrame?.runs[0]?.font.color)).toEqual(['#FFFFFF', '#FFFFFF'])
    expect(await applyFill('offWhite')).toBe('Off-white fill, navy text, on 2 shapes.')
    expect(a.textFrame?.runs[0]?.font.color).toBe('#022338')
    expect(await applyFill('none')).toBe('No fill on 2 shapes.')
    expect(a.fill.type).toBe('NoFill')
    expect(a.textFrame?.runs[0]?.font.color).toBe('#022338')
  })

  it('opens groups up into their members, where PowerPoint allows', async () => {
    const m1 = shape({ id: 'm1' })
    const m2 = shape({ id: 'm2' })
    deck([group([m1, group([m2], { id: 'inner' })], { id: 'g' })], ['g'])
    expect(await applyFill('green')).toBe('Green fill, white text, on 2 shapes.')
    expect([m1, m2].map((m) => m.fill.foregroundColor)).toEqual(['#328840', '#328840'])
    deck([group([shape(), shape()], { id: 'g' })], ['g'], '1.7')
    await expect(applyFill('green')).rejects.toThrow('Select shapes or text boxes to fill.')
  })

  it('colours text, with a reminder about white', async () => {
    const a = shape({ id: 'a', text: 'Text' })
    deck([a, line({ id: 'l' })], ['a', 'l'])
    expect(await applyTextColour('charcoal')).toBe('Charcoal text in 1 shape.')
    expect(a.textFrame?.runs[0]?.font.color).toBe('#4A4A4C')
    expect(await applyTextColour('white')).toBe('White text in 1 shape. Use white on navy or green only.')
  })

  it('outlines in light grey, or none, leaving lines alone', async () => {
    const a = shape({ id: 'a' })
    const p = picture({ id: 'p' })
    const l = line({ id: 'l' })
    deck([a, p, l], ['a', 'p', 'l'])
    expect(await applyOutline('grey')).toBe('Light grey outline on 2 shapes.')
    expect([a.line, p.line].map((x) => [x.visible, x.color, x.weight])).toEqual([[true, '#E8E9E6', 0.75], [true, '#E8E9E6', 0.75]])
    expect(l.line.color).toBe('#156082')
    await applyOutline('none')
    expect(a.line.visible).toBe(false)
  })

  it('asks for something to format', async () => {
    deck([shape({ id: 'a' })], [])
    await expect(applyFill('navy')).rejects.toThrow('Select the shapes to format first.')
    deck([picture({ id: 'p' })], ['p'])
    await expect(applyTextColour('navy')).rejects.toThrow(/Pictures, lines and tables have no text frame/)
  })
})

describe('fonts and text frames', () => {
  it('sets Candara on title placeholders and Aptos on everything else', async () => {
    const title = placeholder('Title', 'A title', { id: 't', font: { name: 'Calibri' } })
    const body = placeholder('Body', 'Text', { id: 'b', font: { name: 'Arial' } })
    const box = textBox('Note', { id: 'x', font: { name: 'Times New Roman' } })
    deck([title, body, box], ['t', 'b', 'x'])
    expect(await applyBwbFonts()).toBe('Candara on 1 title, Aptos on 2 shapes.')
    expect([title, body, box].map((s) => s.textFrame?.runs[0]?.font.name)).toEqual(['Candara', 'Aptos', 'Aptos'])
  })

  it('finds titles by name where PowerPoint cannot report placeholder types', async () => {
    const title = placeholder('Title', 'A title', { id: 't', name: 'Title 1', font: { name: 'Calibri' } })
    const sub = placeholder('Subtitle', 'Sub', { id: 's', name: 'Subtitle 2', font: { name: 'Calibri' } })
    deck([title, sub], ['t', 's'], '1.7')
    await applyBwbFonts()
    expect([title, sub].map((s) => s.textFrame?.runs[0]?.font.name)).toEqual(['Candara', 'Aptos'])
  })

  it('sets Aptos in table cells with PowerPointApi 1.9', async () => {
    const t = tableShape([['A', 'B'], ['1', '2']], { id: 't' })
    deck([t], ['t'])
    expect(await applyBwbFonts()).toBe('Aptos in 1 table.')
    expect(Object.keys(t.table?.cells ?? {})).toEqual(['0,0', '0,1', '1,0', '1,1'])
    expect(t.table?.cells['1,1']).toEqual({ font: { name: 'Aptos' } })
  })

  it('sets text margins, autofit, wrapping and the vertical anchor', async () => {
    const a = shape({ id: 'a', text: 'x' })
    deck([a, picture({ id: 'p' })], ['a', 'p'])
    expect(await setTextMargins('zero')).toBe('No text margins in 1 shape.')
    expect(a.textFrame).toMatchObject({ leftMargin: 0, rightMargin: 0, topMargin: 0, bottomMargin: 0 })
    await setTextMargins('standard')
    expect(a.textFrame?.leftMargin).toBeCloseTo(5.67, 2)
    expect(STANDARD_MARGIN).toBeCloseTo(72 * 0.2 / 2.54, 10)
    a.textFrame && (a.textFrame.autoSizeSetting = 'AutoSizeShapeToFitText')
    expect(await setAutofitOff()).toBe('Autofit off for 1 shape.')
    expect(a.textFrame?.autoSizeSetting).toBe('AutoSizeNone')
    await setWordWrap(false)
    expect(a.textFrame?.wordWrap).toBe(false)
    expect(await setVerticalAnchor('Bottom')).toBe('Text anchored at the bottom of 1 shape.')
    expect(a.textFrame?.verticalAlignment).toBe('Bottom')
  })
})

describe('cleaning up wording', () => {
  it('finds fixes paragraph by paragraph, with offsets in the whole text', () => {
    const text = 'The program \u2014 a pilot\rIt said "yes"'
    const fixes = wordingFixes(text)
    for (const f of fixes) expect(f.end).toBeGreaterThan(f.start)
    expect(fixes.map((f) => text.slice(f.start, f.end))).toEqual(expect.arrayContaining(['"', ' \u2014 ']))
    expect(wordingFixes('')).toEqual([])
  })

  it('fixes house style in the selected shapes and keeps the formatting around each fix', async () => {
    const frame = textFrame('', { runs: [run('The organization '), run('must', { bold: true }), run(' act \u2014 now. He said "go".')] })
    const t = textBox('', { id: 't', textFrame: frame })
    deck([t], ['t'])
    const message = await cleanSelectedText()
    expect(message).toMatch(/^\d+ changes made\.$/)
    expect(textOf(t)).toBe('The organisation must act, now. He said “go”.')
    expect(t.textFrame?.runs.find((r) => r.text === 'must')?.font.bold).toBe(true)
  })

  it('cleans every paragraph and every selected shape', async () => {
    const a = textBox('Colour is fine.\nThe color is not.', { id: 'a' })
    const b = shape({ id: 'b', text: 'We analyze it.' })
    deck([a, b], ['a', 'b'])
    expect(await cleanSelectedText()).toBe('2 changes made.')
    expect([textOf(a), textOf(b)]).toEqual(['Colour is fine.\nThe colour is not.', 'We analyse it.'])
  })

  it('says when there is nothing to change', async () => {
    deck([textBox('All in order.', { id: 'a' })], ['a'])
    expect(await cleanSelectedText()).toBe('Nothing to change in the selection.')
  })
})
