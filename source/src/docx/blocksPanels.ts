/** Panels: callouts, key figures, snapshot grids, icon grids and case studies (builder.Doc). */
import type { Doc } from './builder.ts'
import { stripIconTokens } from './iconPolicy.ts'
import { iconPara, iconXml } from './inline.ts'
import type { CaseStudyStat, IconGridItem, SnapshotItem, StatItem } from './types.ts'
import {
  CHARCOAL,
  cpLen,
  GREEN,
  LIME,
  NAVY,
  OFFWHITE,
  para,
  PY_SPACE,
  pyRound,
  pySplit,
  pyStrip,
  run,
  spacer,
  TEXT_W,
  WHITE,
} from './xml.ts'

// "USD 102bn": a currency code set small so the figure never wraps
const CURRENCY = new RegExp(`^([A-Z]{3})[${PY_SPACE}]+([^\\n]+)$`)

function equalWidths(total: number, n: number): number[] {
  const w = Math.floor(total / n)
  const ws = Array<number>(n).fill(w)
  ws[n - 1] = w + total - w * n
  return ws
}

function textList(text: string | readonly string[] | null | undefined): readonly string[] {
  return typeof text === 'string' ? [text] : (text ?? [])
}

export function callout(
  doc: Doc,
  title: string | null | undefined,
  text: string | readonly string[] | null | undefined,
  variant = 'green',
  bulletItems?: readonly string[] | null,
  icon?: string | null,
): void {
  const green = variant === 'green'
  const color = green ? GREEN : CHARCOAL
  const style = green ? 'BwBHighlight' : 'BwBHighlightCharcoal'
  let inner = ''
  if (title || icon) {
    let head = ''
    if (icon) {
      head = iconXml(doc, icon, green ? 'green' : 'charcoal', 0.6, { fontPt: 11 })
      if (head && title) head += run('\u{2002}')
    }
    inner += para(head + (title ? doc.runs(title) : ''), { style: 'BwBCalloutTitle' })
  }
  for (const t of textList(text)) inner += para(doc.runs(t), { style: 'BwBCalloutText', spacing: { after: 80 } })
  for (const b of bulletItems ?? []) inner += para(doc.runs(b), { style: 'BwBBulletGreenCircle' })
  if (!inner) inner = para('', { style: 'BwBCalloutText' })
  const borders = `<w:tcBorders><w:left w:val="single" w:sz="32" w:space="0" w:color="${color}"/></w:tcBorders>`
  const mar =
    '<w:tcMar><w:top w:w="240" w:type="dxa"/><w:left w:w="320" w:type="dxa"/>' +
    '<w:bottom w:w="200" w:type="dxa"/><w:right w:w="280" w:type="dxa"/></w:tcMar>'
  const row = '<w:tr><w:trPr><w:cantSplit/></w:trPr>' + doc.tc(TEXT_W, inner, OFFWHITE, borders, mar) + '</w:tr>'
  doc.blockTable(doc.tblWrap(style, [TEXT_W], row, '', '04A0', title || 'Callout'))
  doc.add(spacer(160))
}

/** Navy panel with two to four headline numbers; `hero` is the index shown in lime. */
export function stats(doc: Doc, items: readonly StatItem[], hero?: number | null): void {
  const n = items.length
  if (!n) {
    doc.warn('A key figures block has no items; it was skipped.')
    return
  }
  const ws = equalWidths(TEXT_W, n)
  const w = ws[0] ?? 0
  const mar =
    '<w:tcMar><w:top w:w="320" w:type="dxa"/><w:left w:w="280" w:type="dxa"/>' +
    '<w:bottom w:w="280" w:type="dxa"/><w:right w:w="120" w:type="dxa"/></w:tcMar>'
  const innerPt = (w - 400) / 20.0
  const split = items.map((it): [string, string] => {
    const v = pyStrip(String(it.value))
    const m = CURRENCY.exec(v)
    return m ? [m[1] ?? '', m[2] ?? ''] : ['', v]
  })
  let size = 40
  for (const [pre, rest] of split) {
    const widthPerPt = (pre ? 3 * 0.68 * 0.4 + 0.12 : 0) + 0.62 * cpLen(rest.replaceAll('*', ''))
    size = Math.min(size, Math.max(18, Math.trunc(innerPt / Math.max(widthPerPt, 0.1))))
  }
  const isHero = (i: number): boolean => hero !== null && hero !== undefined && i === hero
  let cells = ''
  doc.navy = true
  try {
    items.forEach((it, i) => {
      const st = isHero(i) ? 'BwBStatNumberAccent' : 'BwBStatNumber'
      const [pre, rest] = split[i] ?? ['', '']
      const runs = (pre ? run(pre + ' ', { sz: Math.max(11, pyRound(size * 0.4)) }) : '') + doc.runs(rest, { sz: size, icon_cm: 0.8 })
      const num = para(runs, { style: st, spacing: { before: 0, after: 80, line: Math.trunc(size * 22), lineRule: 'exact' } })
      const ic = it.icon ? iconPara(doc, it.icon, it.icon_colour || (isHero(i) ? 'lime' : 'white'), 0.8, 100, true, true) : ''
      cells += doc.tc(ws[i] ?? 0, ic + num + para(doc.runs(it.label ?? ''), { style: 'BwBStatLabel' }), NAVY, null, mar)
    })
  } finally {
    doc.navy = false
  }
  const row = '<w:tr><w:trPr><w:cantSplit/></w:trPr>' + cells + '</w:tr>'
  const borders = `<w:tblBorders><w:insideV w:val="single" w:sz="4" w:space="0" w:color="${WHITE}"/></w:tblBorders>`
  doc.blockTable(doc.tblWrap('BwBSnapshot', ws, row, borders, '04A0', 'Key figures'))
  doc.add(spacer(160))
}

/** Navy grid of key facts: eyebrow, value and note per cell. */
export function snapshot(doc: Doc, items: readonly SnapshotItem[], columnsIn: unknown = 3): void {
  const columns = Math.max(1, Math.trunc(Number(columnsIn)) || 1)
  const ws = equalWidths(TEXT_W, columns)
  const mar =
    '<w:tcMar><w:top w:w="240" w:type="dxa"/><w:left w:w="280" w:type="dxa"/>' +
    '<w:bottom w:w="240" w:type="dxa"/><w:right w:w="240" w:type="dxa"/></w:tcMar>'
  let rows = ''
  for (let i = 0; i < items.length; i += columns) {
    const chunk = items.slice(i, i + columns)
    let cells = ''
    for (let c = 0; c < columns; c++) {
      const it: SnapshotItem = chunk[c] ?? {}
      let inner = ''
      doc.navy = true
      try {
        if (it.icon) inner += iconPara(doc, it.icon, it.icon_colour || 'white', 0.6, 80, true, true)
        if (it.eyebrow) inner += para(doc.runs(it.eyebrow), { style: 'BwBSnapshotEyebrow' })
        if (it.value) inner += para(doc.runs(it.value), { style: 'BwBSnapshotValue' })
        if (it.note) inner += para(doc.runs(it.note), { style: 'BwBSnapshotNote' })
      } finally {
        doc.navy = false
      }
      cells += doc.tc(ws[c] ?? 0, inner || para('', { style: 'BwBSnapshotNote' }), NAVY, null, mar)
    }
    rows += '<w:tr><w:trPr><w:cantSplit/></w:trPr>' + cells + '</w:tr>'
  }
  const borders =
    `<w:tblBorders><w:insideH w:val="single" w:sz="4" w:space="0" w:color="${WHITE}"/>` +
    `<w:insideV w:val="single" w:sz="4" w:space="0" w:color="${WHITE}"/></w:tblBorders>`
  doc.blockTable(doc.tblWrap('BwBSnapshot', ws, rows, borders, '04A0', 'Snapshot'))
  doc.add(spacer(160))
}

/**
 * Tiles with an icon, a short title and one or two sentences. `light` tiles are off-white with green icons; `navy`
 * tiles have white icons and one lime accent chosen with `hero`.
 */
export function iconGrid(
  doc: Doc,
  itemsIn: readonly IconGridItem[] | null | undefined,
  columnsIn: unknown = 3,
  variant = 'light',
  hero?: number | null,
): void {
  const items = [...(itemsIn ?? [])]
  if (!items.length) return
  const columns = Math.max(1, Math.min(Math.trunc(Number(columnsIn || 3)) || 1, 4, items.length))
  const navy = variant === 'navy'
  const ws = equalWidths(TEXT_W, columns)
  const mar =
    '<w:tcMar><w:top w:w="260" w:type="dxa"/><w:left w:w="260" w:type="dxa"/>' +
    '<w:bottom w:w="240" w:type="dxa"/><w:right w:w="220" w:type="dxa"/></w:tcMar>'
  const fill = navy ? NAVY : OFFWHITE
  let rows = ''
  for (let i = 0; i < items.length; i += columns) {
    const chunk = items.slice(i, i + columns)
    let cells = ''
    for (let c = 0; c < columns; c++) {
      const n = i + c
      const it = chunk[c]
      if (it === undefined || it === null) {
        cells += doc.tc(ws[c] ?? 0, para('', { style: 'BwBCalloutText' }), fill, null, mar)
        continue
      }
      let inner = ''
      doc.navy = navy
      try {
        if (it.icon) {
          const col = it.icon_colour || (navy ? (hero === n ? 'lime' : 'white') : 'green')
          inner += iconPara(doc, it.icon, col, 0.8, 120, true, navy)
        }
        if (it.title) inner += para(doc.runs(it.title), { style: navy ? 'BwBSnapshotValue' : 'BwBCalloutTitle' })
        for (const t of textList(it.text)) {
          inner += para(doc.runs(t), { style: navy ? 'BwBSnapshotNote' : 'BwBCalloutText', spacing: { after: 60 } })
        }
      } finally {
        doc.navy = false
      }
      cells += doc.tc(ws[c] ?? 0, inner || para('', { style: 'BwBCalloutText' }), fill, null, mar)
    }
    rows += '<w:tr><w:trPr><w:cantSplit/></w:trPr>' + cells + '</w:tr>'
  }
  const borders =
    `<w:tblBorders><w:insideH w:val="single" w:sz="24" w:space="0" w:color="${WHITE}"/>` +
    `<w:insideV w:val="single" w:sz="24" w:space="0" w:color="${WHITE}"/></w:tblBorders>`
  doc.blockTable(doc.tblWrap(navy ? 'BwBSnapshot' : 'BwBPlain', ws, rows, borders, '04A0', 'Icon grid'))
  doc.add(spacer(160))
}

/** Navy title panel on the left, text and figures on the right. */
export function caseStudy(
  doc: Doc,
  title: string,
  meta: string | null | undefined,
  body: string | readonly string[],
  statItems?: readonly CaseStudyStat[] | null,
  source?: string | null,
  icon?: string | null,
): void {
  const lw = 3300
  const rw = TEXT_W - 3300
  const words = pySplit(stripIconTokens(title).replace(/\*\*|\^\[[^\n]*?\]/g, ''))
  const longest = words.length ? Math.max(...words.map(cpLen)) : 1
  // the longest word must fit the panel
  const size = Math.max(14, Math.min(22, Math.trunc((lw - 600) / 20.0 / (0.62 * Math.max(longest, 1)))))
  let left = ''
  doc.navy = true
  try {
    left = icon ? iconPara(doc, icon, 'white', 1.0, 160, true, true) : ''
    left += para(doc.runs(title, { sz: size, icon_cm: 0.6 }), {
      style: 'BwBPanelTitle',
      spacing: { before: 0, after: 240, line: Math.trunc(size * 23.6), lineRule: 'exact' },
    })
    if (meta) {
      left += para('', {
        style: 'BwBPanelText',
        spacing: { before: 120, after: 120, line: 20, lineRule: 'exact' },
        extra: `<w:pBdr><w:top w:val="single" w:sz="12" w:space="1" w:color="${LIME}"/></w:pBdr>`,
      })
      left += para(doc.runs(meta), { style: 'BwBPanelText' })
    }
  } finally {
    doc.navy = false
  }
  let right = ''
  for (const t of textList(body)) right += para(doc.runs(t), { style: null, spacing: { after: 200 } })
  if (statItems && statItems.length) {
    const sws = equalWidths(rw - 2 * 200, statItems.length)
    const cells = statItems
      .map(
        (s, i) =>
          doc.tc(
            sws[i] ?? 0,
            para(doc.runs(String(s.value)), { style: 'BwBStatNumberLight' }) +
              para(doc.runs(s.label ?? ''), { style: 'BwBStatLabelLight' }),
          ),
      )
      .join('')
    right += doc.tblWrap('BwBPlain', sws, '<w:tr><w:trPr><w:cantSplit/></w:trPr>' + cells + '</w:tr>')
  }
  right += para(source ? doc.runs(`Source: ${source}`) : '', { style: 'BwBSource', spacing: { before: 160, after: 0 } })
  const marL =
    '<w:tcMar><w:top w:w="360" w:type="dxa"/><w:left w:w="320" w:type="dxa"/><w:bottom w:w="360" w:type="dxa"/>' +
    '<w:right w:w="280" w:type="dxa"/></w:tcMar>'
  const marR =
    '<w:tcMar><w:top w:w="360" w:type="dxa"/><w:left w:w="400" w:type="dxa"/><w:bottom w:w="240" w:type="dxa"/>' +
    '<w:right w:w="0" w:type="dxa"/></w:tcMar>'
  const row =
    '<w:tr><w:trPr><w:cantSplit/></w:trPr>' + doc.tc(lw, left, NAVY, null, marL) + doc.tc(rw, right, null, null, marR) + '</w:tr>'
  doc.blockTable(doc.tblWrap('BwBPlain', [lw, rw], row, '', '04A0', title))
  doc.add(spacer(160))
}
