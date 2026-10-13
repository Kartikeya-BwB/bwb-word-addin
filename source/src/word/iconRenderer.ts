// Draws Lucide icons in BwB colours in the browser, for the document builder (PNG with an SVG copy for Word).
// The SVG comes from the builder's own helpers, which match lucide-static for every icon.

import { loadIconIndex } from '../brand/icons.ts'
import { LUCIDE_ALIASES, createIconResolver, lucideSvgFromNodes, recolourSvg } from '../docx/index.ts'
import type { IconColour, IconRenderer } from '../docx/types.ts'

/** Draws SVG markup to a PNG `px` wide and `pxHigh` high (square by default). */
export async function svgToPng(svg: string, px: number, pxHigh = px): Promise<Uint8Array> {
  // Size the SVG itself, so every browser rasterises it at full resolution rather than scaling up 24 px.
  const sized = pxHigh === px ? svg.replace(/width="\d+" height="\d+"/, `width="${px}" height="${px}"`) : svg
  const image = new Image(px, pxHigh)
  image.src = `data:image/svg+xml;charset=utf-8,${encodeURIComponent(sized)}`
  await image.decode()
  const canvas = document.createElement('canvas')
  canvas.width = px
  canvas.height = pxHigh
  const context = canvas.getContext('2d')
  if (!context) throw new Error('This browser cannot draw icons (no canvas).')
  context.drawImage(image, 0, 0, px, pxHigh)
  const blob = await new Promise<Blob | null>((resolve) => canvas.toBlob(resolve, 'image/png'))
  if (!blob) throw new Error('The icon could not be drawn.')
  return new Uint8Array(await blob.arrayBuffer())
}

export async function createIconRenderer(): Promise<IconRenderer> {
  const index = await loadIconIndex()
  const names = Object.keys(index.icons)
  const resolve = createIconResolver(names, { ...LUCIDE_ALIASES, ...index.aliases })
  const svg = (name: string, colour: IconColour): string | null => {
    const icon = index.icons[name]
    return icon ? recolourSvg(lucideSvgFromNodes(icon.n), colour) : null
  }
  return {
    resolve,
    svg,
    png: async (name: string, colour: IconColour, px: number) => {
      const drawn = svg(name, colour)
      return drawn ? svgToPng(drawn, px) : null
    },
    names: () => [...names, ...Object.keys(index.aliases)],
  }
}
