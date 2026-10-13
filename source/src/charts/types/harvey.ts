// Harvey ball tables: options down the left, criteria across the top, and a ball filled in quarters for each score
// from 0 to 4. Scores can be typed as 0 to 4, as 0% to 100%, or as the ball characters themselves.

import { COLOURS } from '../../brand/tokens.ts'
import type { ChartContext } from '../frame.ts'
import { LABELS_CUT, chartFrame, warn } from '../frame.ts'
import type { Grid } from '../grid.ts'
import { normaliseGrid } from '../grid.ts'
import { fitLabels, lineHeight, textWidth } from '../measure.ts'
import { INK } from '../palette.ts'
import { circle, group, line, num, path, textLines, textLinesCentred } from '../svg.ts'
import { placeholder } from './empty.ts'

const GLYPHS: Record<string, number> = { '\u{25CB}': 0, '\u{25D4}': 1, '\u{25D1}': 2, '\u{25D5}': 3, '\u{25CF}': 4 }

/** A score from 0 to 4, or null for an empty cell or text that is not a score. */
export function harveyScore(input: string): number | null {
  const s = input.trim()
  if (!s) return null
  const glyph = GLYPHS[s]
  if (glyph !== undefined) return glyph
  const percent = /^(\d+(?:\.\d+)?)\s*%$/.exec(s)
  if (percent) return Math.min(4, Math.max(0, Math.round(Number(percent[1]) / 25)))
  const n = Number(s.replace(',', '.'))
  if (!Number.isFinite(n)) return null
  return Math.min(4, Math.max(0, Math.round(n)))
}

/** True when a cell is exactly a score: 0 to 4, a ball character, or a multiple of 25%. */
function exactScore(cell: string): boolean {
  const s = cell.trim()
  if (GLYPHS[s] !== undefined || /^[0-4]$/.test(s)) return true
  const percent = /^(\d+)\s*%$/.exec(s)
  return percent ? Number(percent[1]) % 25 === 0 && Number(percent[1]) <= 100 : false
}

export interface HarveyTable {
  options: string[]
  criteria: string[]
  scores: Array<Array<number | null>>
}

export function harveyTable(input: Grid, onWarning: (message: string) => void = () => {}): HarveyTable {
  const grid = normaliseGrid(input)
  const header = grid[0] ?? []
  const width = header.length
  const criteria = header.slice(1).map((c, i) => c.trim() || `Criterion ${i + 1}`)
  const rows = grid.slice(1).filter((r) => r.some((c) => c.trim()))
  let odd = 0
  const scores = rows.map((r) =>
    Array.from({ length: Math.max(0, width - 1) }, (_, c) => {
      const cell = r[c + 1] ?? ''
      const score = harveyScore(cell)
      if (cell.trim() && !exactScore(cell)) odd += 1
      return score
    }),
  )
  if (odd) onWarning('Harvey ball scores run from 0 to 4; other values are rounded or left blank.')
  return { options: rows.map((r, i) => (r[0] ?? '').trim() || `Option ${i + 1}`), criteria, scores }
}

/** One Harvey ball: an outlined circle filled clockwise from twelve o'clock in quarters. */
export function harveyBall(cx: number, cy: number, r: number, score: number, colour: string): string {
  const stroke = { colour, width: Math.max(0.75, r * 0.12) }
  if (score >= 4) return circle(cx, cy, r, colour, stroke)
  const parts = [circle(cx, cy, r, INK.white, stroke)]
  if (score > 0) {
    const turn = score / 4
    const ex = cx + r * Math.sin(turn * 2 * Math.PI)
    const ey = cy - r * Math.cos(turn * 2 * Math.PI)
    const large = turn > 0.5 ? 1 : 0
    parts.push(path(`M${num(cx)} ${num(cy)}L${num(cx)} ${num(cy - r)}A${num(r)} ${num(r)} 0 ${large} 1 ${num(ex)} ${num(ey)}Z`, colour))
    parts.push(circle(cx, cy, r, 'none', stroke))
  }
  return parts.join('')
}

export function renderHarvey(ctx: ChartContext): string {
  const table = harveyTable(ctx.spec.data, (m) => warn(ctx, m))
  if (!table.options.length || !table.criteria.length) return placeholder(ctx, 'Add options down the first column and criteria across the first row.')
  const font = ctx.font
  const lh = lineHeight(font)
  const frame = chartFrame(ctx, [])
  const plot = frame.plot
  const labelWidth = Math.min(plot.w * 0.34, Math.max(...table.options.map((o) => textWidth(o, font, true))) + font)
  const columnWidth = (plot.w - labelWidth) / table.criteria.length
  const header = fitLabels(table.criteria, Math.max(font, columnWidth - font * 0.6), font, 3)
  const headerH = header.lines * lineHeight(header.size) + font * 0.6
  const rowH = (plot.h - headerH) / table.options.length
  const radius = Math.max(3, Math.min(rowH * 0.34, columnWidth * 0.28, font * 1.05))
  const maxLines = Math.min(3, Math.max(1, Math.floor((rowH * 0.95) / lh)))
  const options = fitLabels(table.options, labelWidth - font, font, maxLines)
  if ([...header.wrapped, ...options.wrapped].some((w) => w.truncated)) warn(ctx, LABELS_CUT)
  const wanted = ctx.spec.highlight?.trim().toLowerCase() ?? ''
  const highlighted = wanted ? table.options.findIndex((o) => o.toLowerCase() === wanted) : -1
  if (wanted && highlighted < 0) warn(ctx, `There is no option called \u{201C}${ctx.spec.highlight ?? ''}\u{201D} to highlight.`)
  const rules: string[] = []
  const balls: string[] = []
  const labels: string[] = []
  const top = plot.y + headerH
  header.wrapped.forEach((w, c) => {
    const cx = plot.x + labelWidth + columnWidth * (c + 0.5)
    const blockTop = top - font * 0.45 - w.lines.length * lineHeight(header.size)
    labels.push(textLines(cx, blockTop, w.lines, { size: header.size, fill: INK.text, anchor: 'middle', bold: true }))
  })
  rules.push(line(plot.x, top, plot.x + plot.w, top, INK.text, 0.75))
  table.options.forEach((_, r) => {
    const cy = top + rowH * (r + 0.5)
    const isHighlight = r === highlighted
    labels.push(textLinesCentred(plot.x, cy, options.wrapped[r]?.lines ?? [], { size: options.size, fill: INK.text, bold: isHighlight }))
    if (r < table.options.length - 1) rules.push(line(plot.x, top + rowH * (r + 1), plot.x + plot.w, top + rowH * (r + 1), INK.grid, 0.5))
    table.criteria.forEach((__, c) => {
      const score = table.scores[r]?.[c] ?? null
      if (score === null) return
      const cx = plot.x + labelWidth + columnWidth * (c + 0.5)
      balls.push(harveyBall(cx, cy, radius, score, isHighlight ? COLOURS.green : COLOURS.navy))
    })
  })
  return frame.svg + group('rules', rules.join('')) + group('balls', balls.join('')) + group('labels', labels.join(''))
}
