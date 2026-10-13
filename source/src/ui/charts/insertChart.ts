// Puts a finished chart in the document. PowerPoint takes the SVG as a vector graphic in the slide's content area.
// Word takes a high-resolution PNG; with a title, the picture and its "Figure N." caption go in together as one BwB
// figure block, centred and kept with its caption.

import { TAGS } from '../../brand/slideGrid.ts'
import type { RenderedChart } from '../../charts/spec.ts'
import type { Host } from '../../host.ts'
import { insertSvgHere } from '../../media/insert.ts'
import { insertBlocks } from '../../word/insert.ts'

export interface InsertOptions {
  host: Host
  title: string
  source: string
  /** The source is drawn in the chart, so the caption leaves it out. */
  sourceInChart: boolean
  left?: number
  top?: number
}

/** Four times 96 dpi, the resolution insertSvgHere uses for Word. */
export function pixelSize(chart: Pick<RenderedChart, 'widthPt' | 'heightPt'>): { width: number; height: number } {
  const width = Math.max(48, Math.round((chart.widthPt / 72) * 96 * 4))
  return { width, height: Math.round((width * chart.heightPt) / chart.widthPt) }
}

/** The SVG with its width and height set in pixels, so the browser draws it at full resolution. */
export function sizedSvg(svg: string, width: number, height: number): string {
  return svg.replace(/<svg\b([^>]*)>/, (_match, attrs: string) => {
    const kept = attrs.replace(/\s(width|height)="[^"]*"/g, '')
    return `<svg${kept} width="${width}" height="${height}">`
  })
}

export async function insertChart(chart: RenderedChart, options: InsertOptions): Promise<string> {
  const { svg, widthPt, heightPt, altText } = chart
  if (options.host === 'powerpoint') {
    await insertSvgHere({
      svg, widthPt, heightPt, altText, tags: { [TAGS.chart]: options.title.trim() || 'chart' },
      ...(options.left !== undefined ? { left: options.left } : {}),
      ...(options.top !== undefined ? { top: options.top } : {}),
    })
    return 'Chart inserted on the slide. Right-click it and choose Convert to Shape to edit its parts.'
  }
  const title = options.title.trim()
  if (!title) {
    await insertSvgHere({ svg, widthPt, heightPt, altText })
    return 'Chart inserted. Add a title to insert it with a numbered caption.'
  }
  // One block for picture and caption: inserting the caption after the picture would replace the picture's
  // paragraph, which Word reports as empty text.
  const { svgToPng } = await import('../../word/iconRenderer.ts')
  const px = pixelSize(chart)
  const png = await svgToPng(sizedSvg(svg, px.width, px.height), px.width, px.height)
  const source = options.sourceInChart ? '' : options.source.trim()
  const { warnings } = await insertBlocks([{
    type: 'figure', data: png, ext: 'png', widthPx: px.width, heightPx: px.height,
    width_cm: Math.round(((widthPt * 2.54) / 72) * 100) / 100, alt: altText, caption: title,
    ...(source ? { source } : {}),
  }])
  const done = 'Chart inserted with its numbered caption.'
  return warnings.length ? `${done} ${warnings.join(' ')}` : done
}
