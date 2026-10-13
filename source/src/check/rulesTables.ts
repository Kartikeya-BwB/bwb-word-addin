// Tables and pictures: BwB table styles, captions below tables, alt text on pictures.

import { DATA_TABLE_STYLE_NAMES, LAYOUT_TABLE_STYLE_NAMES, TABLE_STYLES } from '../brand/styles.ts'
import type { DocSnapshot, Finding, NeighbourParagraph } from './types.ts'
import { excerpt } from './util.ts'

const isCaption = (p: NeighbourParagraph | null, kind: 'Table' | 'Figure'): boolean =>
  !!p && (p.style === `BwB ${kind === 'Table' ? 'Table' : 'Image'} Caption` || new RegExp(`^\\s*${kind}\\s+\\d+`, 'i').test(p.text))

export function tableFindings(snapshot: DocSnapshot): Finding[] {
  const out: Finding[] = []
  for (const t of snapshot.tables) {
    const label = t.firstCell ? excerpt(t.firstCell, 0, 40, 0) : `Table ${t.index + 1}`
    const known = DATA_TABLE_STYLE_NAMES.has(t.style) || LAYOUT_TABLE_STYLE_NAMES.has(t.style)
    if (!known) {
      out.push({
        id: `tables:style:${t.index}`, rule: 'tables/style', group: 'tables', severity: 'minor',
        title: 'Table not in a BwB table style',
        detail: `It uses “${t.style || 'no style'}”. BwB data tables have a navy header row, light grey borders and right-aligned figures.`,
        excerpt: label, location: { kind: 'table', table: t.index },
        fix: { kind: 'table-style', table: t.index, style: TABLE_STYLES.standard.name }, fixLabel: 'Apply BwB Standard',
      })
    }
    const isData = DATA_TABLE_STYLE_NAMES.has(t.style) || !known
    // An abbreviations table is a list, not data: it needs no caption.
    const isAbbreviations = /^(?:abbreviations?|acronyms?)$/i.test(t.firstCell.trim())
    if (isData && !isAbbreviations && t.rows > 1 && !isCaption(t.after, 'Table')) {
      const above = isCaption(t.before, 'Table')
      out.push({
        id: `tables:caption:${t.index}`, rule: above ? 'tables/caption-above' : 'tables/no-caption', group: 'tables',
        severity: above ? 'copyedit' : 'minor',
        title: above ? 'Caption above the table' : 'Table without a caption',
        detail: above
          ? 'In the BwB format the caption sits below the table: “Table 1. Title. Source: ...”.'
          : 'Number and source every data table with a caption below it (Insert, Caption).',
        excerpt: label, location: { kind: 'table', table: t.index },
      })
    }
  }
  for (const p of snapshot.pictures) {
    if (!p.altTitle.trim() && !p.altDescription.trim()) {
      out.push({
        id: `tables:alt:${p.index}`, rule: 'tables/alt-text', group: 'tables', severity: 'minor',
        title: 'Picture without alt text', detail: 'Describe what the picture shows (right-click, View Alt Text).',
        location: { kind: 'picture', picture: p.index },
      })
    }
  }
  return out
}
