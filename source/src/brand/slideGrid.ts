// The BwB slide grid for 16:9 slides (13.333 by 7.5 inches, 960 by 540 points). The slide template, the alignment
// tools and the deck check all use these numbers, in points, so shapes snap to the same lines everywhere.

export interface Area {
  left: number
  top: number
  width: number
  height: number
}

export const SLIDE: Area = { left: 0, top: 0, width: 960, height: 540 }

export const MARGINS = { left: 40, right: 40, top: 28, bottom: 36 } as const

/** Where every content slide's action title sits: Candara bold 24 pt navy, two lines at most. */
export const TITLE_AREA: Area = { left: 40, top: 28, width: 880, height: 64 }

/** Everything between the title and the source line. */
export const CONTENT_AREA: Area = { left: 40, top: 108, width: 880, height: 360 }

/** The source line and footnotes: Aptos 8 pt charcoal. */
export const SOURCE_AREA: Area = { left: 40, top: 476, width: 700, height: 24 }

/** The footer band: logo at the left, slide number at the right, the colour strip along the bottom edge. */
export const FOOTER_AREA: Area = { left: 40, top: 508, width: 880, height: 22 }
export const STRIP_HEIGHT = 4

/** Inside the margins: the area "align to margins" uses. */
export const SAFE_AREA: Area = {
  left: MARGINS.left,
  top: MARGINS.top,
  width: SLIDE.width - MARGINS.left - MARGINS.right,
  height: SLIDE.height - MARGINS.top - MARGINS.bottom,
}

export const COLUMNS = 12
export const GUTTER = 16
export const COLUMN_WIDTH = (SAFE_AREA.width - (COLUMNS - 1) * GUTTER) / COLUMNS

/** Left and right edges of each of the 12 columns, in points. */
export function columnEdges(): Array<{ left: number; right: number }> {
  return Array.from({ length: COLUMNS }, (_, i) => {
    const left = SAFE_AREA.left + i * (COLUMN_WIDTH + GUTTER)
    return { left, right: left + COLUMN_WIDTH }
  })
}

/** Slide type sizes, in points. Nothing on a slide is set below 8 pt, and body text stays at 10 pt or more. */
export const SLIDE_TYPE = {
  title: { font: 'Candara', size: 24, bold: true, colour: '022338' },
  subtitle: { font: 'Aptos', size: 14, bold: false, colour: '4A4A4C' },
  body: { font: 'Aptos', size: 14, bold: false, colour: '000000' },
  small: { font: 'Aptos', size: 11, bold: false, colour: '000000' },
  source: { font: 'Aptos', size: 8, bold: false, colour: '4A4A4C' },
  minBody: 10,
  minAny: 8,
} as const

/** Tags the add-in writes on shapes and slides, so its own elements can be found, updated and removed. */
export const TAGS = {
  element: 'BWB_ELEMENT',
  sticker: 'BWB_STICKER',
  note: 'BWB_NOTE',
  tracker: 'BWB_TRACKER',
  section: 'BWB_SECTION',
  layout: 'BWB_LAYOUT',
  rag: 'BWB_RAG',
  chart: 'BWB_CHART',
} as const
