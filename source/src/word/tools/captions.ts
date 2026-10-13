// Figure captions under pictures, and alt text taken from them. A caption is a BwB caption block inserted after the
// picture's paragraph (insertBlocks would replace a picture's line, which has no text, so this inserts it itself).
// Each caption is numbered by its place among the captions above it, so the number is right before fields update.

import type { Block } from '../docxApi.ts'
import { buildBlockOoxml } from '../docxApi.ts'
import { isSeqField, updateFields } from '../fields.ts'
import { UserFacingError, runWord, supports } from '../run.ts'
import { captionKind, captionTitle, count } from './common.ts'
import type { PictureInfo } from './pictures.ts'
import { allPictures, isFigurePicture, selectedPictures } from './pictures.ts'

const PLACEHOLDER = { text: '{field:Title}', source: '{field:Source}' }

interface Placed {
  info: PictureInfo
  /** The paragraph after the picture's line, if any. */
  next: Word.Paragraph
  /** Figure captions above the picture. */
  before: number
  /** True when the picture shares its line with the picture before it. */
  sameLine: boolean
}

/** Where each picture sits: what follows it, how many figure captions come before it, whether it shares a line. */
async function place(context: Word.RequestContext, list: PictureInfo[]): Promise<Placed[]> {
  const paragraphs = context.document.body.paragraphs
  paragraphs.load('items/text,items/style,items/styleBuiltIn')
  const next = list.map((p) => p.paragraph.getNextOrNullObject())
  next.forEach((p) => p.load('text,style,styleBuiltIn'))
  await context.sync()
  const captions = paragraphs.items.filter((p) => captionKind(p) === 'Figure').map((p) => p.getRange('Whole'))
  const compare = list.map((p) => {
    const at = p.paragraph.getRange('Whole')
    return captions.map((c) => c.compareLocationWith(at))
  })
  const previous = list.map((p, i) => (i > 0 ? list[i - 1]?.paragraph.getRange('Whole').compareLocationWith(p.paragraph.getRange('Whole')) : null))
  await context.sync()
  return list.map((info, i) => ({
    info,
    next: next[i] ?? info.paragraph.getNextOrNullObject(),
    before: (compare[i] ?? []).filter((r) => r.value === 'Before' || r.value === 'AdjacentBefore').length,
    sameLine: previous[i]?.value === 'Equal',
  }))
}

const hasCaption = (p: Placed): boolean => !p.next.isNullObject && captionKind(p.next) === 'Figure'

async function captionXml(number: number, text: string, source?: string): Promise<{ ooxml: string; warnings: string[] }> {
  const block: Block = { type: 'caption', kind: 'Figure', text, ...(source ? { source } : {}) }
  return buildBlockOoxml([block], { seqStart: { Figure: number } })
}

/** Renumbers the caption fields after a caption goes in above others. Word on the web may not update fields. */
async function renumber(): Promise<void> {
  if (!supports('1.5')) return
  try {
    await updateFields(isSeqField)
  } catch {
    // The numbers shown are already right for the new captions; F9 refreshes the rest.
  }
}

/** Adds a numbered figure caption below the selected picture (or the picture on the cursor's line). */
export async function captionSelectedPicture(title = '', source = ''): Promise<string> {
  const result = await runWord(async (context) => {
    const [first] = await selectedPictures(context)
    if (!first) throw new UserFacingError('Select a picture, or put the cursor on its line, first.')
    if (first.inTable) throw new UserFacingError('This picture is in a table: put the caption below the table instead.')
    const [placed] = await place(context, [first])
    if (!placed) throw new UserFacingError('Select a picture, or put the cursor on its line, first.')
    if (hasCaption(placed)) throw new UserFacingError('This picture already has a caption below it.')
    const number = placed.before + 1
    const text = title.trim() || PLACEHOLDER.text
    const { ooxml, warnings } = await captionXml(number, text, title.trim() ? source.trim() : PLACEHOLDER.source)
    first.paragraph.getRange('Whole').insertOoxml(ooxml, 'After').select('End')
    await context.sync()
    return { number, warnings, placeholder: !title.trim() }
  })
  await renumber()
  const tip = result.placeholder ? ' Type over the grey placeholders.' : ''
  return [`Figure ${result.number} caption added below the picture.${tip}`, ...result.warnings].join(' ')
}

/** Adds a caption with placeholders under every figure that has none, in one pass. */
export async function captionAllPictures(): Promise<string> {
  const result = await runWord(async (context) => {
    const figures = (await allPictures(context)).filter(isFigurePicture)
    if (!figures.length) throw new UserFacingError('There are no figures in this document. Icons and logos are not captioned.')
    const placed = (await place(context, figures)).filter((p) => !p.sameLine)
    const missing = placed.filter((p) => !hasCaption(p))
    let added = 0
    for (const p of missing) {
      const { ooxml } = await captionXml(p.before + added + 1, PLACEHOLDER.text, PLACEHOLDER.source)
      p.info.paragraph.getRange('Whole').insertOoxml(ooxml, 'After')
      added += 1
    }
    await context.sync()
    return { added, had: placed.length - missing.length }
  })
  if (!result.added) return 'Every figure already has a caption.'
  await renumber()
  const had = result.had ? ` ${count(result.had, 'figure')} already had one.` : ''
  return `Captions added under ${count(result.added, 'figure')}. Type over the grey placeholders.${had}`
}

/** Fills empty alt text from the caption under each picture: the selected pictures, or every picture. */
export async function altTextFromCaptions(): Promise<string> {
  return runWord(async (context) => {
    const chosen = await selectedPictures(context, false)
    const list = chosen.length ? chosen : await allPictures(context)
    if (!list.length) throw new UserFacingError('There are no pictures in this document.')
    const empty = list.filter((p) => !p.alt)
    if (!empty.length) return chosen.length ? 'The selected pictures already have alt text.' : 'Every picture already has alt text.'
    const placed = await place(context, empty)
    let filled = 0
    for (const p of placed) {
      const title = hasCaption(p) ? captionTitle(p.next.text) : ''
      if (!title) continue
      p.info.picture.altTextDescription = title
      filled += 1
    }
    await context.sync()
    const left = empty.length - filled
    const rest = left ? ` ${count(left, 'picture has', 'pictures have')} no caption to use: describe ${left === 1 ? 'it' : 'them'} by hand (right-click, View Alt Text).` : ''
    return filled ? `Alt text filled in for ${count(filled, 'picture')} from ${filled === 1 ? 'its caption' : 'their captions'}.${rest}` : `No alt text filled in.${rest}`
  })
}
