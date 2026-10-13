import { CONTENT_AREA, FOOTER_AREA, SAFE_AREA, SLIDE, SOURCE_AREA, STRIP_HEIGHT, TITLE_AREA, columnEdges } from '../../brand/slideGrid.ts'
import type { Area } from '../../brand/slideGrid.ts'

const LABELS: Array<{ area: Area; label: string; tone: string }> = [
  { area: TITLE_AREA, label: 'Action title, Candara bold 24 pt', tone: 'var(--navy)' },
  { area: CONTENT_AREA, label: 'Content area', tone: 'var(--green)' },
  { area: SOURCE_AREA, label: 'Source, 8 pt', tone: 'var(--charcoal)' },
  { area: FOOTER_AREA, label: 'Logo · slide number', tone: 'var(--charcoal)' },
]

/** The BwB slide grid at a glance: the 12 columns, the margins and where titles, content and sources go. */
export function SlideGridDiagram() {
  const cm = (pt: number): string => `${((pt * 2.54) / 72).toFixed(1)} cm`
  return (
    <figure className="grid-diagram">
      <svg viewBox={`0 0 ${SLIDE.width} ${SLIDE.height}`} role="img" aria-label="The BwB slide grid">
        <rect width={SLIDE.width} height={SLIDE.height} fill="var(--white)" stroke="var(--charcoal-20)" />
        {columnEdges().map((c, i) => (
          <rect key={i} x={c.left} y={SAFE_AREA.top} width={c.right - c.left} height={SAFE_AREA.height} fill="var(--green-20)" />
        ))}
        {LABELS.map(({ area, label, tone }) => (
          <g key={label}>
            <rect x={area.left} y={area.top} width={area.width} height={area.height} fill="none" stroke={tone} strokeWidth={2} strokeDasharray="8 5" />
            <text x={area.left + 8} y={area.top + Math.min(area.height - 6, 22)} fontSize={Math.min(18, area.height - 6)} fill={tone}>{label}</text>
          </g>
        ))}
        <rect x={0} y={SLIDE.height - STRIP_HEIGHT} width={SLIDE.width * 0.5} height={STRIP_HEIGHT} fill="var(--green)" />
        <rect x={SLIDE.width * 0.5} y={SLIDE.height - STRIP_HEIGHT} width={SLIDE.width * 0.33} height={STRIP_HEIGHT} fill="var(--charcoal)" />
        <rect x={SLIDE.width * 0.83} y={SLIDE.height - STRIP_HEIGHT} width={SLIDE.width * 0.17} height={STRIP_HEIGHT} fill="var(--lime)" />
      </svg>
      <figcaption className="field-hint">
        16:9 slide, {cm(SAFE_AREA.left)} margins, 12 columns with {cm(16)} gutters. Arrange tools snap to these lines.
      </figcaption>
    </figure>
  )
}
