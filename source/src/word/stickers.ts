// Status stickers ("Draft", "Confidential" and so on) in the page header, the way consulting toolbars mark
// slides. Each sticker is a tagged content control, so it can be found, replaced and removed reliably.

import { runWord } from './run.ts'

export const STICKERS = ['Draft', 'For discussion', 'Preliminary', 'Illustrative', 'Confidential', 'Not for distribution'] as const
export type Sticker = (typeof STICKERS)[number]

export const STICKER_TAG = 'BwBSticker'
const HEADER_TYPES = ['Primary', 'FirstPage', 'EvenPages'] as const

interface HeaderInfo {
  body: Word.Body
  stickers: Word.ContentControlCollection
  pictures: Word.InlinePictureCollection
}

async function loadHeaders(context: Word.RequestContext): Promise<HeaderInfo[]> {
  const sections = context.document.sections
  sections.load('items')
  await context.sync()
  const headers: HeaderInfo[] = []
  for (const section of sections.items) {
    for (const type of HEADER_TYPES) {
      const body = section.getHeader(type)
      const stickers = body.contentControls
      const pictures = body.inlinePictures
      body.load('text')
      stickers.load('items/tag')
      pictures.load('items')
      headers.push({ body, stickers, pictures })
    }
  }
  await context.sync()
  return headers
}

/** Replaces any sticker in the document's headers with `label`, or removes them all when `label` is null. */
export async function setSticker(label: Sticker | null): Promise<number> {
  return runWord(async (context) => {
    const headers = await loadHeaders(context)
    for (const h of headers) {
      h.stickers.items.filter((c) => c.tag === STICKER_TAG).forEach((c) => c.delete(false))
    }
    await context.sync()
    if (!label) return 0

    // Prefer headers that already carry content (the BwB logo header), so covers and dividers stay clean.
    const used = headers.filter((h) => h.body.text.trim() !== '' || h.pictures.items.length > 0)
    const primaryOnly = headers.filter((_, i) => i % HEADER_TYPES.length === 0)
    const targets = used.length ? used : primaryOnly
    let added = 0
    for (const h of targets) {
      // Linked headers are the same part; skip any that already received a sticker in this run.
      const existing = h.body.contentControls
      existing.load('items/tag')
      await context.sync()
      if (existing.items.some((c) => c.tag === STICKER_TAG)) continue
      const paragraph = h.body.insertParagraph(label.toUpperCase(), 'End')
      paragraph.alignment = 'Right'
      paragraph.spaceAfter = 0
      paragraph.spaceBefore = 0
      paragraph.font.set({ name: 'Aptos', size: 8, bold: true, color: '#4A4A4C', italic: false })
      const control = paragraph.insertContentControl()
      control.tag = STICKER_TAG
      control.title = 'BwB sticker'
      control.appearance = 'Hidden'
      added += 1
    }
    await context.sync()
    return added
  })
}

/** The sticker currently in the headers, if any. */
export async function currentSticker(): Promise<string | null> {
  return runWord(async (context) => {
    const headers = await loadHeaders(context)
    for (const h of headers) {
      const sticker = h.stickers.items.find((c) => c.tag === STICKER_TAG)
      if (sticker) {
        sticker.load('text')
        await context.sync()
        return sticker.text.trim() || null
      }
    }
    return null
  })
}
