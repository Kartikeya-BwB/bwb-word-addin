// Picture into frame: fills the selected shape with a photo from the BwB library, cropped to the shape's proportions
// so PowerPoint does not stretch it. The crop is made on a canvas at about twice the shape's size on screen, at most
// 2000 pixels along the longer side. Picture fills need PowerPointApi 1.8.

import type { ImageEntry } from '../../brand/images.ts'
import { fetchImageBytes } from '../../brand/images.ts'
import { toBase64 } from '../../media/base64.ts'
import { UserFacingError } from '../../word/run.ts'
import { requirePpt, runPpt, supportsPpt } from '../run.ts'

export type Focus = 'top' | 'centre' | 'bottom'

export interface Crop {
  x: number
  y: number
  width: number
  height: number
}

/**
 * The largest part of the picture with the frame's proportions. Too wide, it keeps the middle; too tall, it keeps the
 * top, the middle or the bottom, as asked (faces tend to sit in the top half).
 */
export function cropFor(imageWidth: number, imageHeight: number, frameWidth: number, frameHeight: number, focus: Focus = 'centre'): Crop {
  const frame = frameWidth / frameHeight
  if (imageWidth / imageHeight > frame) {
    const width = imageHeight * frame
    return { x: (imageWidth - width) / 2, y: 0, width, height: imageHeight }
  }
  const height = imageWidth / frame
  const spare = imageHeight - height
  return { x: 0, y: focus === 'top' ? 0 : focus === 'bottom' ? spare : spare / 2, width: imageWidth, height }
}

/** Twice the shape's size at 96 pixels an inch, at most `max` pixels along the longer side, and never above the crop. */
export function outputSize(frameWidth: number, frameHeight: number, crop: Crop, max = 2000): { width: number; height: number } {
  let width = (frameWidth / 72) * 96 * 2
  let height = (frameHeight / 72) * 96 * 2
  const shrink = Math.min(1, max / Math.max(width, height), crop.width / width, crop.height / height)
  width *= shrink
  height *= shrink
  return { width: Math.max(1, Math.round(width)), height: Math.max(1, Math.round(height)) }
}

/** Turns photo bytes into JPEG bytes cropped for a frame of the given size in points. */
export type Renderer = (bytes: Uint8Array, frame: { width: number; height: number }, focus: Focus) => Promise<Uint8Array>

/** The browser renderer: decodes the photo, crops and scales it on a canvas, and encodes a JPEG. */
export const canvasRenderer: Renderer = async (bytes, frame, focus) => {
  const url = URL.createObjectURL(new Blob([bytes.slice()]))
  try {
    const image = new Image()
    image.src = url
    await image.decode()
    const crop = cropFor(image.naturalWidth, image.naturalHeight, frame.width, frame.height, focus)
    const size = outputSize(frame.width, frame.height, crop)
    const canvas = document.createElement('canvas')
    canvas.width = size.width
    canvas.height = size.height
    const context = canvas.getContext('2d')
    if (!context) throw new UserFacingError('This browser cannot crop pictures.')
    context.imageSmoothingQuality = 'high'
    context.drawImage(image, crop.x, crop.y, crop.width, crop.height, 0, 0, size.width, size.height)
    const blob = await new Promise<Blob | null>((resolve) => canvas.toBlob(resolve, 'image/jpeg', 0.88))
    if (!blob) throw new UserFacingError('The photo could not be prepared.')
    return new Uint8Array(await blob.arrayBuffer())
  } finally {
    URL.revokeObjectURL(url)
  }
}

/** Shapes a picture fill suits: plain geometric shapes and text boxes. */
const FILLABLE = new Set(['GeometricShape', 'TextBox'])

export interface FrameDeps {
  fetchBytes: (image: ImageEntry) => Promise<Uint8Array>
  render: Renderer
}

/** Fills the one selected shape with the photo, cropped to fit, and gives it the photo's title as alt text (1.10). */
export async function fillWithPicture(image: ImageEntry, focus: Focus = 'centre', deps: FrameDeps = { fetchBytes: fetchImageBytes, render: canvasRenderer }): Promise<string> {
  requirePpt('1.8', 'Filling a shape with a photo')
  const target = await runPpt(async (context) => {
    const shapes = context.presentation.getSelectedShapes()
    shapes.load('items/id,items/name,items/type,items/width,items/height')
    await context.sync()
    const [shape] = shapes.items
    if (shapes.items.length !== 1 || !shape || !FILLABLE.has(shape.type)) {
      throw new UserFacingError('Select one shape, such as a rectangle or a circle, to fill with the photo.')
    }
    if (shape.width < 1 || shape.height < 1) throw new UserFacingError('The shape is too small to hold a photo.')
    return { id: shape.id, name: shape.name, width: shape.width, height: shape.height }
  })
  const jpeg = await deps.render(await deps.fetchBytes(image), target, focus)
  const described = supportsPpt('1.10')
  await runPpt(async (context) => {
    const shape = context.presentation.getSelectedShapes().getItemOrNullObject(target.id)
    shape.load('id')
    await context.sync()
    if (shape.isNullObject) throw new UserFacingError(`“${target.name}” is no longer selected. Select it and try again.`)
    shape.fill.setImage(toBase64(jpeg))
    if (described) shape.altTextDescription = image.title
    await context.sync()
  })
  return `“${target.name}” filled with “${image.title}”.`
}
