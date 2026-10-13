// Picture tools: the full text width or an exact width (keeping proportions), the width of the first, alignment.
// They act on the selected pictures; with none selected, on every figure: a picture on its own line outside a
// table, at least 3 cm wide and not an icon or logo, so the BwB marks in a document are left as they are.

import { UserFacingError, runWord } from '../run.ts'
import { TEXT_WIDTH_CM, cmToPt, count, formatCm, isBlank, ptToCm } from './common.ts'

const MIN_FIGURE_CM = 3
const ICON_OR_LOGO = /\blogo\b|^\s*icon\b|\bBwB icon\b/i
const PICTURE_PROPS = 'items/width,items/height,items/altTextTitle,items/altTextDescription'

export interface PictureInfo {
  picture: Word.InlinePicture
  paragraph: Word.Paragraph
  width: number
  height: number
  /** On its own line: the paragraph has no text. */
  alone: boolean
  inTable: boolean
  alt: string
}

export const isFigurePicture = (p: PictureInfo): boolean =>
  p.alone && !p.inTable && ptToCm(p.width) >= MIN_FIGURE_CM && !ICON_OR_LOGO.test(p.alt)

async function describe(context: Word.RequestContext, pictures: Word.InlinePicture[]): Promise<PictureInfo[]> {
  const paragraphs = pictures.map((p) => p.paragraph)
  paragraphs.forEach((p) => p.load('text,tableNestingLevel'))
  await context.sync()
  return pictures.map((picture, i) => {
    const paragraph = paragraphs[i] ?? picture.paragraph
    return {
      picture,
      paragraph,
      width: picture.width,
      height: picture.height,
      alone: isBlank(paragraph.text),
      inTable: paragraph.tableNestingLevel > 0,
      alt: `${picture.altTextTitle} ${picture.altTextDescription}`.trim(),
    }
  })
}

/**
 * The pictures in the selection. Unless `cursorLine` is false, a cursor on a picture's own line (a paragraph with no
 * text) counts as choosing it; an icon inside a line of text has to be selected itself.
 */
export async function selectedPictures(context: Word.RequestContext, cursorLine = true): Promise<PictureInfo[]> {
  const selection = context.document.getSelection()
  const direct = selection.inlinePictures
  direct.load(PICTURE_PROPS)
  const paragraphs = selection.paragraphs
  paragraphs.load('items/text')
  await context.sync()
  if (direct.items.length || !cursorLine) return describe(context, direct.items)
  const onTheirLine = paragraphs.items.filter((p) => isBlank(p.text)).map((p) => p.inlinePictures)
  onTheirLine.forEach((c) => c.load(PICTURE_PROPS))
  await context.sync()
  return describe(context, onTheirLine.flatMap((c) => c.items))
}

/** Every picture in the body, in document order. */
export async function allPictures(context: Word.RequestContext): Promise<PictureInfo[]> {
  const all = context.document.body.inlinePictures
  all.load(PICTURE_PROPS)
  await context.sync()
  return describe(context, all.items)
}

/** The selected pictures, or every figure when none is selected. */
async function targets(context: Word.RequestContext): Promise<{ list: PictureInfo[]; chosen: boolean; others: number }> {
  const chosen = await selectedPictures(context)
  if (chosen.length) return { list: chosen, chosen: true, others: 0 }
  const all = await allPictures(context)
  const figures = all.filter(isFigurePicture)
  return { list: figures, chosen: false, others: all.length - figures.length }
}

function resize(p: PictureInfo, widthPt: number): void {
  const height = p.width > 0 ? (p.height * widthPt) / p.width : p.height
  p.picture.width = widthPt
  p.picture.height = height
}

const NO_FIGURES = 'There are no figures to change: select a picture first. Icons, logos and pictures in tables are left alone unless selected.'

/**
 * Sets pictures to a width in centimetres, keeping their proportions; 'text' is the full BwB text width (16.6 cm).
 * Fitting to the text width leaves pictures in tables alone, since they would overflow their cells.
 */
export async function setPictureWidth(width: number | 'text'): Promise<string> {
  const cm = width === 'text' ? TEXT_WIDTH_CM : width
  if (!Number.isFinite(cm) || cm < 0.5 || cm > 50) throw new UserFacingError('Give a width between 0.5 and 50 cm.')
  return runWord(async (context) => {
    const { list, chosen, others } = await targets(context)
    const inTables = width === 'text' && chosen ? list.filter((p) => p.inTable).length : 0
    const changing = width === 'text' ? list.filter((p) => !p.inTable) : list
    if (!changing.length) {
      throw new UserFacingError(chosen ? 'Pictures in tables are not fitted to the text width: it would overflow the cell.' : NO_FIGURES)
    }
    changing.forEach((p) => resize(p, cmToPt(cm)))
    await context.sync()
    const done = `${count(changing.length, 'picture')} set to ${formatCm(cm)} cm wide, keeping ${changing.length === 1 ? 'its' : 'their'} proportions.`
    const note = chosen
      ? inTables ? ` ${count(inTables, 'picture in a table was', 'pictures in tables were')} left alone.` : ''
      : others ? ` ${count(others, 'other picture was', 'other pictures were')} left alone: icons, logos, small pictures and pictures in tables or text.` : ''
    return `${chosen ? '' : 'No picture was selected, so every figure was changed: '}${done}${note}`
  })
}

/** Gives every selected picture the width of the first, keeping proportions. */
export async function matchPictureWidths(): Promise<string> {
  return runWord(async (context) => {
    const list = await selectedPictures(context)
    const [first, ...rest] = list
    if (!first || !rest.length) throw new UserFacingError('Select two or more pictures. The first one sets the width.')
    rest.forEach((p) => resize(p, first.width))
    await context.sync()
    return `${count(rest.length, 'picture')} set to ${formatCm(ptToCm(first.width))} cm wide, the width of the first.`
  })
}

export type PictureAlignment = 'Left' | 'Centered' | 'Right'

const ALIGNED: Record<PictureAlignment, string> = { Left: 'aligned left', Centered: 'centred', Right: 'aligned right' }

/** Aligns the lines that hold the selected pictures, or every figure when none is selected. */
export async function alignPictures(alignment: PictureAlignment): Promise<string> {
  return runWord(async (context) => {
    const { list, chosen } = await targets(context)
    if (!list.length) throw new UserFacingError(NO_FIGURES)
    list.forEach((p) => {
      p.paragraph.alignment = alignment
    })
    await context.sync()
    return `${chosen ? '' : 'No picture was selected, so every figure was changed: '}${count(list.length, 'picture')} ${ALIGNED[alignment]}.`
  })
}
