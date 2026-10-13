import { afterEach, describe, expect, it } from 'vitest'
import { STICKER_TAG, currentSticker, setSticker } from '../../src/word/stickers.ts'
import { createModel, para, section } from '../fakeWord/model.ts'
import { installFakeWord, uninstallFakeWord } from '../fakeWord/install.ts'

afterEach(uninstallFakeWord)

describe('stickers', () => {
  it('puts the sticker in the header that carries the logo, as a hidden tagged control', async () => {
    const model = createModel([para('Body')])
    installFakeWord(model)
    expect(await setSticker('Draft')).toBe(1)
    const header = model.sections[0]?.headers.Primary
    const control = header?.controls.find((c) => c.tag === STICKER_TAG)
    expect(control?.appearance).toBe('Hidden')
    expect(control?.paragraphs[0]?.text).toBe('DRAFT')
    expect(control?.paragraphs[0]?.alignment).toBe('Right')
    expect(control?.paragraphs[0]?.font).toMatchObject({ name: 'Aptos', size: 8, bold: true, color: '#4A4A4C' })
    expect(await currentSticker()).toBe('DRAFT')
  })

  it('replaces one sticker with another rather than stacking them', async () => {
    const model = createModel([para('Body')])
    installFakeWord(model)
    await setSticker('Draft')
    await setSticker('For discussion')
    const controls = model.sections[0]?.headers.Primary.controls.filter((c) => c.tag === STICKER_TAG) ?? []
    expect(controls).toHaveLength(1)
    expect(await currentSticker()).toBe('FOR DISCUSSION')
  })

  it('removes every sticker, text included', async () => {
    const model = createModel([para('Body')])
    installFakeWord(model)
    await setSticker('Confidential')
    expect(await setSticker(null)).toBe(0)
    const header = model.sections[0]?.headers.Primary
    expect(header?.controls).toHaveLength(0)
    expect(header?.paragraphs.some((p) => p.text === 'CONFIDENTIAL')).toBe(false)
    expect(await currentSticker()).toBeNull()
  })

  it('uses the primary header of each section when no header has content yet', async () => {
    const model = createModel([para('Body')], { sections: [section(), section()] })
    installFakeWord(model)
    expect(await setSticker('Preliminary')).toBe(2)
    expect(model.sections.every((s) => s.headers.FirstPage.controls.length === 0)).toBe(true)
  })
})
