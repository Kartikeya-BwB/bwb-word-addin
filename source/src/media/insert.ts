// Inserting pictures (photos, icons, charts) in whichever app the add-in is running in. Word places them inline at
// the cursor; PowerPoint places them on the current slide, centred unless a position is given.

import { currentHost } from '../host.ts'
import { hasPowerPoint, supportsPpt } from '../ppt/run.ts'
import { insertPicture } from '../word/marks.ts'
import { toBase64 } from './base64.ts'
import { UserFacingError } from '../word/run.ts'

const PT_PER_CM = 72 / 2.54

export interface PictureRequest {
  /** PNG or JPEG bytes. */
  bytes: Uint8Array
  widthPt: number
  heightPt: number
  altText: string
  /** PowerPoint only: top-left corner on the slide, in points. */
  left?: number
  top?: number
  /** PowerPoint only: tags for the new shape, such as BWB_CHART, so the add-in can find it again. */
  tags?: Record<string, string>
}

/** A 16:9 slide in points, the size of the BwB slide template. */
export const SLIDE = { width: 960, height: 540 } as const

function setSelectedImage(base64: string, options: Office.SetSelectedDataOptions): Promise<void> {
  return new Promise((resolve, reject) => {
    Office.context.document.setSelectedDataAsync(base64, { coercionType: Office.CoercionType.Image, ...options }, (result) => {
      if (result.status === Office.AsyncResultStatus.Succeeded) resolve()
      else reject(new UserFacingError(`PowerPoint did not insert the picture (${result.error.message}).`))
    })
  })
}

/**
 * PowerPoint selects what setSelectedDataAsync placed, but the call takes no alt text. Give the new shape its alt
 * text and tags, so the deck check does not flag the add-in's own pictures. The picture is in either way: if this
 * fails, the author can still add alt text by hand.
 */
async function labelPlacedShape(altText: string, tags: Record<string, string> = {}, at: { left: number; top: number }): Promise<void> {
  if (!hasPowerPoint() || !supportsPpt('1.5')) return
  try {
    await PowerPoint.run(async (context) => {
      const shapes = context.presentation.getSelectedShapes()
      shapes.load('items/id,items/left,items/top')
      await context.sync()
      const [shape] = shapes.items
      // Only the shape just placed: never relabel something else that happens to be selected.
      if (!shape || shapes.items.length !== 1 || Math.abs(shape.left - at.left) > 2 || Math.abs(shape.top - at.top) > 2) return
      if (supportsPpt('1.10')) shape.altTextDescription = altText
      for (const [key, value] of Object.entries(tags)) shape.tags.add(key, value)
      await context.sync()
    })
  } catch {
    // Alt text is a courtesy here; the insert itself succeeded.
  }
}

export async function insertPictureHere(request: PictureRequest): Promise<void> {
  if (currentHost() === 'word') {
    await insertPicture(request.bytes, request.widthPt / PT_PER_CM, request.heightPt / PT_PER_CM, request.altText)
    return
  }
  if (typeof Office === 'undefined') throw new UserFacingError('Open this in PowerPoint to use it.')
  const at = { left: request.left ?? (SLIDE.width - request.widthPt) / 2, top: request.top ?? (SLIDE.height - request.heightPt) / 2 }
  await setSelectedImage(toBase64(request.bytes), {
    imageWidth: request.widthPt,
    imageHeight: request.heightPt,
    imageLeft: at.left,
    imageTop: at.top,
  })
  await labelPlacedShape(request.altText, request.tags, at)
}

export const cmToPt = (cm: number): number => cm * PT_PER_CM

export interface SvgRequest {
  svg: string
  widthPt: number
  heightPt: number
  altText: string
  left?: number
  top?: number
  tags?: Record<string, string>
}

/**
 * Inserts an SVG graphic: in PowerPoint as a vector graphic people can recolour or convert to shapes; in Word as a
 * picture drawn at four times 96 dpi, because Word on the web cannot take SVG.
 */
export async function insertSvgHere(request: SvgRequest): Promise<void> {
  if (currentHost() === 'word') {
    const { svgToPng } = await import('../word/iconRenderer.ts')
    const px = Math.max(48, Math.round((request.widthPt / 72) * 96 * 4))
    const sized = request.svg.replace(/<svg\b([^>]*)>/, (_match, attrs: string) => {
      const kept = attrs.replace(/\s(width|height)="[^"]*"/g, '')
      return `<svg${kept} width="${px}" height="${Math.round((px * request.heightPt) / request.widthPt)}">`
    })
    const png = await svgToPng(sized, px, Math.round((px * request.heightPt) / request.widthPt))
    await insertPicture(png, request.widthPt / PT_PER_CM, request.heightPt / PT_PER_CM, request.altText)
    return
  }
  if (typeof Office === 'undefined') throw new UserFacingError('Open this in PowerPoint to use it.')
  const at = { left: request.left ?? (SLIDE.width - request.widthPt) / 2, top: request.top ?? (SLIDE.height - request.heightPt) / 2 }
  await new Promise<void>((resolve, reject) => {
    Office.context.document.setSelectedDataAsync(request.svg, {
      coercionType: Office.CoercionType.XmlSvg,
      imageWidth: request.widthPt,
      imageHeight: request.heightPt,
      imageLeft: at.left,
      imageTop: at.top,
    }, (result) => {
      if (result.status === Office.AsyncResultStatus.Succeeded) resolve()
      else reject(new UserFacingError(`PowerPoint did not insert the graphic (${result.error.message}).`))
    })
  })
  await labelPlacedShape(request.altText, request.tags, at)
}
