// What a chart shows when its data has nothing to draw.

import type { ChartContext } from '../frame.ts'
import { wrapText } from '../measure.ts'
import { INK } from '../palette.ts'
import { textLinesCentred } from '../svg.ts'

export function placeholder(ctx: ChartContext, message: string): string {
  const lines = wrapText(message, ctx.width * 0.8, ctx.font, 3).lines
  return textLinesCentred(ctx.width / 2, ctx.height / 2, lines, { size: ctx.font, fill: INK.text, anchor: 'middle' })
}
