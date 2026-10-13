// Inserting pictures (photos, icons, charts) in whichever app the add-in is running in. Word places them inline at
// the cursor; PowerPoint places them on the current slide, centred unless a position is given.

import { currentHost } from '../host.ts'
import { insertPicture } from '../word/marks.ts'
import { toBase64 } from '../word/newDocument.ts'
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

export async function insertPictureHere(request: PictureRequest): Promise<void> {
  if (currentHost() === 'word') {
    await insertPicture(request.bytes, request.widthPt / PT_PER_CM, request.heightPt / PT_PER_CM, request.altText)
    return
  }
  if (typeof Office === 'undefined') throw new UserFacingError('Open this in PowerPoint to use it.')
  await setSelectedImage(toBase64(request.bytes), {
    imageWidth: request.widthPt,
    imageHeight: request.heightPt,
    imageLeft: request.left ?? (SLIDE.width - request.widthPt) / 2,
    imageTop: request.top ?? (SLIDE.height - request.heightPt) / 2,
  })
}

export const cmToPt = (cm: number): number => cm * PT_PER_CM

export interface SvgRequest {
  svg: string
  widthPt: number
  heightPt: number
  altText: string
  left?: number
  top?: number
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
  await new Promise<void>((resolve, reject) => {
    Office.context.document.setSelectedDataAsync(request.svg, {
      coercionType: Office.CoercionType.XmlSvg,
      imageWidth: request.widthPt,
      imageHeight: request.heightPt,
      imageLeft: request.left ?? (SLIDE.width - request.widthPt) / 2,
      imageTop: request.top ?? (SLIDE.height - request.heightPt) / 2,
    }, (result) => {
      if (result.status === Office.AsyncResultStatus.Succeeded) resolve()
      else reject(new UserFacingError(`PowerPoint did not insert the graphic (${result.error.message}).`))
    })
  })
}
