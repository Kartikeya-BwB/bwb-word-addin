import { CONTENT_AREA, FOOTER_AREA, SAFE_AREA, SLIDE, SOURCE_AREA, STRIP_HEIGHT, TITLE_AREA, columnEdges } from '../../brand/slideGrid.ts'
import type { Area } from '../../brand/slideGrid.ts'

interface Zone {
  area: Area
  label: string
  /** The words drawn inside the area; small areas carry only the legend. */
  inside?: string
  tone: string
}

const ZONES: Zone[] = [
  { area: TITLE_AREA, label: 'Action title: Candara bold 24 pt, navy', inside: 'Action title', tone: 'var(--navy)' },
  { area: CONTENT_AREA, label: 'Content: on the 12 columns, inside the margins', inside: 'Content', tone: 'var(--green)' },
  { area: SOURCE_AREA, label: 'Source line: Aptos 8 pt, charcoal', tone: 'var(--charcoal)' },
  { area: FOOTER_AREA, label: 'Footer: logo at the left, slide number at the right', tone: 'var(--navy)' },
]

/** The BwB slide grid at a glance: the 12 columns, the margins and where titles, content and sources go. */
export function SlideGridDiagram() {
  const cm = (pt: number): string => `${((pt * 2.54) / 72).toFixed(1)} cm`
  return (
    <figure className="grid-diagram">
      <svg viewBox={`0 0 ${SLIDE.width} ${SLIDE.height}`} role="img" aria-label="The BwB slide grid">
        <rect width={SLIDE.width} height={SLIDE.height} fill="var(--white)" stroke="var(--charcoal-20)" strokeWidth={2} />
        {columnEdges().map((c, i) => (
          <rect key={i} x={c.left} y={SAFE_AREA.top} width={c.right - c.left} height={SAFE_AREA.height} fill="var(--green-20)" />
        ))}
        {ZONES.map(({ area, label, inside, tone }) => (
          <g key={label}>
            <rect x={area.left} y={area.top} width={area.width} height={area.height} fill="none" stroke={tone} strokeWidth={4} strokeDasharray="14 8" />
            {inside && (
              <text x={area.left + 16} y={area.top + Math.min(area.height - 14, 44)} fontSize={36} fontWeight={700} fill={tone}>{inside}</text>
            )}
          </g>
        ))}
        <rect x={0} y={SLIDE.height - STRIP_HEIGHT * 2} width={SLIDE.width * 0.5} height={STRIP_HEIGHT * 2} fill="var(--green)" />
        <rect x={SLIDE.width * 0.5} y={SLIDE.height - STRIP_HEIGHT * 2} width={SLIDE.width * 0.33} height={STRIP_HEIGHT * 2} fill="var(--charcoal)" />
        <rect x={SLIDE.width * 0.83} y={SLIDE.height - STRIP_HEIGHT * 2} width={SLIDE.width * 0.17} height={STRIP_HEIGHT * 2} fill="var(--lime)" />
      </svg>
      <ul className="grid-legend">
        {ZONES.map(({ label, tone }) => (
          <li key={label}><span className="grid-key" style={{ borderColor: tone }} aria-hidden />{label}</li>
        ))}
      </ul>
      <figcaption className="field-hint">
        16:9 slide, {cm(SAFE_AREA.left)} margins, 12 columns with {cm(16)} gutters. Snap on the Arrange tab lines edges up with the columns.
      </figcaption>
    </figure>
  )
}
