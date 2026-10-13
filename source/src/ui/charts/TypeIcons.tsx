// Small pictures of each chart type for the type picker, drawn in the current text colour on a 24 by 24 grid.

import type { ReactNode } from 'react'
import type { ChartType } from '../../charts/spec.ts'

const LIGHT = 0.45

const bars = (items: Array<[number, number, number, number, number?]>): ReactNode =>
  items.map(([x, y, w, h, opacity], i) => <rect key={i} x={x} y={y} width={w} height={h} fillOpacity={opacity ?? 1} />)

const SHAPES: Record<ChartType, ReactNode> = {
  column: bars([[3, 13, 4, 8], [10, 8, 4, 13], [17, 4, 4, 17]]),
  stackedColumn: bars([[3, 14, 4, 7], [3, 10, 4, 4, LIGHT], [10, 11, 4, 10], [10, 6, 4, 5, LIGHT], [17, 8, 4, 13], [17, 3, 4, 5, LIGHT]]),
  column100: bars([[3, 11, 4, 10], [3, 3, 4, 8, LIGHT], [10, 8, 4, 13], [10, 3, 4, 5, LIGHT], [17, 14, 4, 7], [17, 3, 4, 11, LIGHT]]),
  bar: bars([[3, 3, 18, 4], [3, 10, 13, 4], [3, 17, 8, 4]]),
  stackedBar: bars([[3, 3, 11, 4], [14, 3, 6, 4, LIGHT], [3, 10, 8, 4], [11, 10, 5, 4, LIGHT], [3, 17, 5, 4], [8, 17, 3, 4, LIGHT]]),
  bar100: bars([[3, 3, 12, 4], [15, 3, 6, 4, LIGHT], [3, 10, 8, 4], [11, 10, 10, 4, LIGHT], [3, 17, 15, 4], [18, 17, 3, 4, LIGHT]]),
  line: (
    <>
      <polyline points="3,17 8,12 13,14 21,5" fill="none" stroke="currentColor" strokeWidth="2" strokeLinejoin="round" strokeLinecap="round" />
      <circle cx="8" cy="12" r="1.8" />
      <circle cx="13" cy="14" r="1.8" />
    </>
  ),
  waterfall: bars([[2, 9, 4, 12], [7, 5, 4, 4, LIGHT], [12, 5, 4, 6, LIGHT], [17, 11, 4, 10]]),
  mekko: bars([[2, 3, 9, 9], [2, 12, 9, 9, LIGHT], [12, 3, 6, 12], [12, 15, 6, 6, LIGHT], [19, 3, 3, 6], [19, 9, 3, 12, LIGHT]]),
  gantt: (
    <>
      {bars([[3, 4, 9, 3], [8, 10, 10, 3]])}
      <path d="M17 15.5 L19.5 18 L17 20.5 L14.5 18 Z" />
    </>
  ),
  donut: (
    <>
      <path d="M12 3 A9 9 0 0 1 20.6 14.7 L16.3 13.4 A4.5 4.5 0 0 0 12 7.5 Z" />
      <path d="M20.6 14.7 A9 9 0 1 1 12 3 L12 7.5 A4.5 4.5 0 1 0 16.3 13.4 Z" fillOpacity={LIGHT} />
    </>
  ),
  harvey: (
    <>
      <circle cx="12" cy="12" r="8" fill="none" stroke="currentColor" strokeWidth="1.8" />
      <path d="M12 12 L12 4 A8 8 0 1 1 4 12 Z" />
    </>
  ),
}

export function TypeIcon({ type, size = 24 }: { type: ChartType; size?: number }) {
  return (
    <svg width={size} height={size} viewBox="0 0 24 24" fill="currentColor" aria-hidden focusable="false">
      {SHAPES[type]}
    </svg>
  )
}
