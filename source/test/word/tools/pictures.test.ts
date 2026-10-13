import { afterEach, describe, expect, it } from 'vitest'
import { altTextFromCaptions, captionAllPictures, captionSelectedPicture } from '../../../src/word/tools/captions.ts'
import { CM_TO_PT } from '../../../src/word/tools/common.ts'
import { alignPictures, matchPictureWidths, setPictureWidth } from '../../../src/word/tools/pictures.ts'
import type { FakeModel, FakeParagraph, FakePicture } from '../../fakeWord/model.ts'
import { createModel, para } from '../../fakeWord/model.ts'
import { installFakeWord, uninstallFakeWord } from '../../fakeWord/install.ts'

afterEach(uninstallFakeWord)

const cm = (pt: number | undefined): number => Math.round(((pt ?? 0) / CM_TO_PT) * 100) / 100

function picture(paragraph: FakeParagraph, widthCm: number, alt = '', heightCm = widthCm / 2): FakePicture {
  return { altTextTitle: '', altTextDescription: alt, width: widthCm * CM_TO_PT, height: heightCm * CM_TO_PT, paragraph }
}

/** A document with a chart, an icon in a line of text, a logo, a small picture and a picture in a table. */
function documentWithPictures(): { model: FakeModel; chart: FakePicture; others: FakePicture[]; line: FakeParagraph } {
  const line = para('')
  const text = para('An icon in a line of text.')
  const logoLine = para('')
  const smallLine = para('')
  const cell = para('', 'Normal', { tableNestingLevel: 1 })
  const model = createModel([para('Intro'), line, text, logoLine, smallLine, cell, para('End')])
  const chart = picture(line, 10)
  const others = [
    picture(text, 0.4, 'Icon: leaf'),
    picture(logoLine, 5.3, 'Bankers without Boundaries logo'),
    picture(smallLine, 2),
    picture(cell, 6),
  ]
  model.pictures.push(chart, ...others)
  return { model, chart, others, line }
}

describe('setPictureWidth', () => {
  it('fits every figure to the text width when nothing is selected, keeping proportions and leaving marks alone', async () => {
    const { model, chart, others } = documentWithPictures()
    installFakeWord(model)
    const message = await setPictureWidth('text')
    expect(cm(chart.width)).toBe(16.6)
    expect(cm(chart.height)).toBe(8.3)
    expect(others.map((p) => cm(p.width))).toEqual([0.4, 5.3, 2, 6])
    expect(message).toBe(
      'No picture was selected, so every figure was changed: 1 picture set to 16.6 cm wide, keeping its proportions. ' +
      '4 other pictures were left alone: icons, logos, small pictures and pictures in tables or text.',
    )
  })

  it('changes a selected picture, even a small one, to an exact width', async () => {
    const { model, others } = documentWithPictures()
    const small = others[2]
    if (!small) throw new Error('no small picture')
    model.selectedPictures = [small]
    installFakeWord(model)
    expect(await setPictureWidth(8)).toBe('1 picture set to 8 cm wide, keeping its proportions.')
    expect(cm(small.width)).toBe(8)
    expect(cm(small.height)).toBe(4)
  })

  it('treats the cursor on a picture\u2019s own line as choosing it', async () => {
    const { model, chart, line } = documentWithPictures()
    model.selection = [line]
    installFakeWord(model)
    await setPictureWidth(12)
    expect(cm(chart.width)).toBe(12)
  })

  it('does not fit a selected picture in a table to the text width, and checks the width given', async () => {
    const { model, others } = documentWithPictures()
    model.selectedPictures = others.slice(3)
    installFakeWord(model)
    await expect(setPictureWidth('text')).rejects.toThrow(/would overflow the cell/)
    await expect(setPictureWidth(0)).rejects.toThrow('Give a width between 0.5 and 50 cm.')
  })

  it('explains when there is no figure to change', async () => {
    installFakeWord(createModel([para('Text only')]))
    await expect(setPictureWidth('text')).rejects.toThrow(/There are no figures to change/)
  })
})

describe('matchPictureWidths and alignPictures', () => {
  it('gives the selected pictures the width of the first', async () => {
    const a = para('')
    const b = para('')
    const model = createModel([a, b])
    const first = picture(a, 12)
    const second = picture(b, 6, '', 6)
    model.pictures.push(first, second)
    model.selectedPictures = [first, second]
    installFakeWord(model)
    expect(await matchPictureWidths()).toBe('1 picture set to 12 cm wide, the width of the first.')
    expect([cm(second.width), cm(second.height)]).toEqual([12, 12])
  })

  it('needs two pictures', async () => {
    const { model, chart } = documentWithPictures()
    model.selectedPictures = [chart]
    installFakeWord(model)
    await expect(matchPictureWidths()).rejects.toThrow('Select two or more pictures. The first one sets the width.')
  })

  it('centres the lines holding the figures', async () => {
    const { model, line, others } = documentWithPictures()
    installFakeWord(model)
    expect(await alignPictures('Centered')).toMatch(/1 picture centred\.$/)
    expect(line.alignment).toBe('Centered')
    expect(others.map((p) => p.paragraph?.alignment)).toEqual([undefined, undefined, undefined, undefined])
  })
})

describe('figure captions', () => {
  function captioned(): { model: FakeModel; first: FakeParagraph; second: FakeParagraph; third: FakeParagraph } {
    const first = para('')
    const second = para('')
    const third = para('')
    const model = createModel([
      para('Intro'), first, second, para('Figure 1. Climate finance flows, 2025. Source: OECD (2026)', 'BwB Image Caption'), third, para('End'),
    ])
    model.pictures.push(picture(first, 10), picture(second, 10), picture(third, 10, 'A described map'))
    return { model, first, second, third }
  }

  it('adds a numbered caption below the selected picture, numbered by the captions above it', async () => {
    const { model, third } = captioned()
    model.selection = [third]
    installFakeWord(model)
    const message = await captionSelectedPicture('Adaptation finance by region', 'BwB analysis')
    expect(message).toBe('Figure 2 caption added below the picture.')
    expect(model.inserts[0]?.location).toBe('After')
    expect(model.inserts[0]?.ooxml).toMatch(/>2</)
    const at = model.paragraphs.indexOf(third)
    expect(model.paragraphs[at + 1]?.text).toBe('Figure 2. Adaptation finance by region Source: BwB analysis')
    expect(model.paragraphs[at + 1]?.style).toBe('BwB Image Caption')
  })

  it('uses placeholders when no title is given, and refuses a picture that has a caption', async () => {
    const { model, first, second } = captioned()
    model.selection = [first]
    installFakeWord(model)
    expect(await captionSelectedPicture()).toBe('Figure 1 caption added below the picture. Type over the grey placeholders.')
    model.selection = [second]
    await expect(captionSelectedPicture()).rejects.toThrow('This picture already has a caption below it.')
  })

  it('refuses pictures in tables and explains when no picture is chosen', async () => {
    const { model, others } = documentWithPictures()
    model.selectedPictures = others.slice(3)
    installFakeWord(model)
    await expect(captionSelectedPicture()).rejects.toThrow(/put the caption below the table/)
    model.selectedPictures = []
    model.selection = model.paragraphs.slice(0, 1)
    await expect(captionSelectedPicture()).rejects.toThrow('Select a picture, or put the cursor on its line, first.')
  })

  it('captions every figure that has none in one pass, numbering each by its place', async () => {
    const { model, first, third } = captioned()
    installFakeWord(model)
    expect(await captionAllPictures()).toBe('Captions added under 2 figures. Type over the grey placeholders. 1 figure already had one.')
    const after = (p: FakeParagraph): string => model.paragraphs[model.paragraphs.indexOf(p) + 1]?.text ?? ''
    expect(after(first)).toMatch(/^Figure 1\./)
    expect(after(third)).toMatch(/^Figure 3\./)
    expect(await captionAllPictures()).toBe('Every figure already has a caption.')
  })

  it('gives two pictures on one line one caption', async () => {
    const line = para('')
    const model = createModel([line, para('End')])
    model.pictures.push(picture(line, 7), picture(line, 7))
    installFakeWord(model)
    expect(await captionAllPictures()).toMatch(/^Captions added under 1 figure\./)
  })
})

describe('altTextFromCaptions', () => {
  it('fills empty alt text from the caption under each picture, leaving alt text that is there', async () => {
    const withCaption = para('')
    const placeholder = para('')
    const bare = para('')
    const described = para('')
    const model = createModel([
      withCaption, para('Figure 2. Climate finance flows, 2025. Source: OECD (2026)', 'BwB Image Caption'),
      placeholder, para('Figure 3. [Title]', 'BwB Image Caption'), bare, para('Body text.'), described, para('Figure 4. Map', 'BwB Image Caption'),
    ])
    const pics = [picture(withCaption, 10), picture(placeholder, 10), picture(bare, 10), picture(described, 10, 'A map of Kenya')]
    model.pictures.push(...pics)
    installFakeWord(model)
    expect(await altTextFromCaptions()).toBe(
      'Alt text filled in for 1 picture from its caption. 2 pictures have no caption to use: describe them by hand (right-click, View Alt Text).',
    )
    expect(pics.map((p) => p.altTextDescription)).toEqual(['Climate finance flows, 2025.', '', '', 'A map of Kenya'])
  })
})
