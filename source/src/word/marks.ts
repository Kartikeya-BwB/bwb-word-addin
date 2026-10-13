// Small insertions at the cursor: Harvey balls, page breaks, pictures (icons and logos).

import { runWord } from './run.ts'
import { toBase64 } from './newDocument.ts'

export const HARVEY_BALLS = [
  { glyph: '○', label: 'Empty (0 per cent)' },
  { glyph: '◔', label: 'Quarter (25 per cent)' },
  { glyph: '◑', label: 'Half (50 per cent)' },
  { glyph: '◕', label: 'Three quarters (75 per cent)' },
  { glyph: '●', label: 'Full (100 per cent)' },
] as const

/** Inserts a Harvey ball at the cursor, in navy so it reads as a BwB mark. */
export async function insertHarveyBall(glyph: string): Promise<void> {
  await runWord(async (context) => {
    const range = context.document.getSelection().insertText(glyph, 'Replace')
    range.font.color = '#022338'
    range.select('End')
    await context.sync()
  })
}

export async function insertPageBreak(): Promise<void> {
  await runWord(async (context) => {
    context.document.getSelection().insertBreak('Page', 'After')
    await context.sync()
  })
}

const CM_TO_PT = 72 / 2.54

/** Inserts an inline picture at the cursor at a size in centimetres, with alt text. */
export async function insertPicture(png: Uint8Array, widthCm: number, heightCm: number, altText: string): Promise<void> {
  const base64 = toBase64(png)
  await runWord(async (context) => {
    const picture = context.document.getSelection().insertInlinePictureFromBase64(base64, 'Replace')
    picture.width = widthCm * CM_TO_PT
    picture.height = heightCm * CM_TO_PT
    picture.altTextTitle = altText
    picture.altTextDescription = altText
    picture.getRange('Whole').select('End')
    await context.sync()
  })
}

export async function fetchPng(url: string): Promise<{ bytes: Uint8Array; width: number; height: number }> {
  const response = await fetch(url)
  if (!response.ok) throw new Error(`Could not load ${url} (${response.status}).`)
  const bytes = new Uint8Array(await response.arrayBuffer())
  // PNG stores width and height as big-endian integers at bytes 16 to 23.
  const view = new DataView(bytes.buffer, bytes.byteOffset, bytes.byteLength)
  return { bytes, width: view.getUint32(16), height: view.getUint32(20) }
}

/** Colours the selected text. The Brand tab offers only the colours BwB allows for text on white. */
export async function colourSelection(hex: string): Promise<void> {
  await runWord(async (context) => {
    context.document.getSelection().font.color = `#${hex}`
    await context.sync()
  })
}
