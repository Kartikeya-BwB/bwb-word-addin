/** Data tables and the details table (Doc.table, Doc._weights, Doc._cell_paras, Doc.key_values). */
import type { Doc } from './builder.ts'
import { caption, paragraph } from './blocksText.ts'
import type { RunOpts } from './inline.ts'
import type { ColumnAlign, KeyValueRows, TableBlock, TableCell } from './types.ts'
import {
  cpLen,
  esc,
  LIGHT,
  NAVY,
  OFFWHITE,
  para,
  PY_SPACE,
  pyMax,
  pyReSplitWs,
  pyStripChars,
  pySum,
  spacer,
  TEXT_W,
  type ParaProps,
  type PyNumber,
} from './xml.ts'

const TABLE_STYLES = new Map([
  ['standard', 'BwBStandard'],
  ['comparison', 'BwBComparison'],
  ['status', 'BwBCountryStatus'],
])
const CELL_MARGINS = new Map<string, readonly [number, number]>([
  ['standard', [140, 200]],
  ['comparison', [160, 200]],
  ['status', [140, 200]],
])

// Python's n ** 0.75 for n = 0 to 70 (glibc pow). V8's Math.pow differs in the last bit for 11 of these, and a column
// width is an int() of the weight, so the exact values are kept.
const POW_075: readonly number[] = [
  0.0, 1.0, 1.681792830507429, 2.2795070569547775, 2.8284271247461903, 3.34370152488211, 3.833658625477635,
  4.303517070658851, 4.756828460010884, 5.196152422706632, 5.623413251903491, 6.0401053545372365, 6.4474195909412515,
  6.846325042023046, 7.237624155400388, 7.621991222319221, 8.0, 8.37214402859269, 8.738851890731821,
  9.10049888342792, 9.457416090031758, 9.809897532292203, 10.158205880770257, 10.502577066424818, 10.843224043318138,
  11.180339887498949, 11.514100370997832, 11.844666116572432, 12.172184414459759, 12.496790764308276,
  12.818610191887021, 13.137758379865671, 13.454342644059432, 13.76846278041797, 14.08021180326277,
  14.38967659151819, 14.696938456699069, 15.002073644034486, 15.305153776189938, 15.606246247497747,
  15.905414575341013, 16.202718714300598, 16.498215337821545, 16.791958091449192, 17.083997821097824,
  17.374382779324037, 17.66315881216501, 17.950369528754806, 18.23605645563822, 18.520259177452136,
  18.80301546543197, 19.08436139501884, 19.364331453687083, 19.64295863997752, 19.92027455460579, 20.19630948441476,
  20.471092479852693, 20.744651426583015, 21.017013111765113, 21.288203285487665, 21.558246717785053,
  21.827167251622598, 22.094987852196983, 22.361730652863333, 22.627416997969522, 22.892067482851036,
  23.155701991215324, 23.41833973012293, 23.679999262753405, 23.94069853912678, 24.200454924935872,
]

const PARAGRAPH_BREAK = new RegExp(`\\n[${PY_SPACE}]*\\n`)

function isBullets(v: TableCell | undefined): v is { bullets: string[] } {
  return typeof v === 'object' && v !== null && 'bullets' in v
}

/** A cell as Python's str() gives it (None is "None" when column weights are measured). */
function cellText(v: TableCell | undefined): string {
  if (typeof v === 'string') return v
  if (isBullets(v)) return (v.bullets ?? []).join(' ')
  return v === null || v === undefined ? 'None' : String(v)
}

/** Relative column widths from the contents (Doc._weights). */
export function columnWeights(header: readonly TableCell[] | null, rows: readonly (readonly TableCell[])[], ncols: number): number[] {
  const w: PyNumber[] = []
  for (let c = 0; c < ncols; c++) {
    const vals: (TableCell | undefined)[] = header && header.length && c < header.length ? [header[c]] : []
    for (const r of rows) if (c < r.length) vals.push(r[c])
    let m: PyNumber = { v: 4, isInt: true }
    for (const v of vals) {
      const t = cellText(v)
      const longest = Math.max(...pyReSplitWs(t).map(cpLen))
      m = pyMax(m, { v: POW_075[Math.min(cpLen(t), 70)] ?? 0, isInt: false }, { v: longest, isInt: true })
    }
    w.push(m)
  }
  const tot = pySum(w)
  const shares = w.map((wi) => pyMax({ v: wi.v / tot.v, isInt: false }, { v: 0.09, isInt: false }))
  const s = pySum(shares)
  return shares.map((wi) => wi.v / s.v)
}

function cellParas(doc: Doc, val: TableCell | undefined, style: string, jc: string | null, keepNext = false): string {
  const kw: ParaProps = {}
  if (jc) kw.jc = jc
  if (keepNext) kw.keep_next = true
  const base: RunOpts = doc.cellSz ? { sz: doc.cellSz } : {}
  if (isBullets(val)) {
    return val.bullets
      .map((b) => para(doc.runs(b, base), { style: 'BwBBulletGreenCircle', ...(keepNext ? { keep_next: true } : {}) }))
      .join('')
  }
  const text = val === null || val === undefined ? '' : String(val)
  return text
    .split(PARAGRAPH_BREAK)
    .map((t) => para(doc.runs(pyStripChars(t, '\n'), base), { style, ...kw }))
    .join('')
}

function alignOf(align: readonly ColumnAlign[], c: number): ColumnAlign {
  return align[c] ?? null
}

export function table(doc: Doc, b: TableBlock): void {
  const style = b.style ?? 'standard'
  const tstyle = TABLE_STYLES.get(style) ?? 'BwBStandard'
  const rows = b.rows ?? []
  let header: TableCell[] | null = b.header ?? null
  const ncols = Math.max(header?.length ?? 0, ...rows.map((r) => r.length))
  if (!ncols) {
    doc.warn('A table with no header and no rows was skipped.')
    return
  }
  doc.cellSz = ncols <= 5 ? null : ncols <= 8 ? 9 : 8
  if (ncols > 9) doc.warn(`Table with ${ncols} columns is very wide; consider splitting it or turning it into a figure.`)
  if (header && header.length && header.length < ncols) header = [...header, ...Array<string>(ncols - header.length).fill('')]
  let fr: number[]
  if (b.widths && b.widths.length && b.widths.length !== ncols) {
    doc.warn(`Table widths give ${b.widths.length} columns but the table has ${ncols}; the widths were measured from the contents.`)
  }
  if (b.widths && b.widths.length === ncols && ncols) {
    const tot = pySum(b.widths.map((x) => ({ v: x, isInt: Number.isInteger(x) }))).v
    fr = b.widths.map((x) => x / tot)
  } else {
    fr = columnWeights(header, rows, ncols)
  }
  const ws = fr.map((f) => Math.trunc(TEXT_W * f))
  ws[ws.length - 1] = (ws[ws.length - 1] ?? 0) + TEXT_W - ws.reduce((a, x) => a + x, 0)
  const align: readonly ColumnAlign[] = b.align && b.align.length ? b.align : Array<ColumnAlign>(ncols).fill(null)
  const [mt, ml] = CELL_MARGINS.get(style) ?? [140, 200]
  const tcmar =
    `<w:tcMar><w:top w:w="${mt}" w:type="dxa"/><w:left w:w="${ml}" w:type="dxa"/>` +
    `<w:bottom w:w="${mt}" w:type="dxa"/><w:right w:w="${ml}" w:type="dxa"/></w:tcMar>`
  const bord =
    '<w:tcBorders>' +
    ['top', 'left', 'bottom', 'right'].map((s) => `<w:${s} w:val="single" w:sz="4" w:space="0" w:color="${LIGHT}"/>`).join('') +
    '</w:tcBorders>'
  const trs: string[] = []
  const together = rows.length <= 10 // short tables are kept on one page
  if (header && header.length) {
    let cells = ''
    for (let c = 0; c < ncols; c++) {
      const a = alignOf(align, c)
      const jc = a === 'right' || a === 'center' ? a : null
      cells += doc.tc(ws[c] ?? 0, cellParas(doc, header[c], 'BwBTableHeader', jc), NAVY, bord, tcmar)
    }
    trs.push('<w:tr><w:trPr><w:cantSplit/><w:tblHeader/></w:trPr>' + cells + '</w:tr>')
  }
  rows.forEach((r, ri) => {
    const last = ri === rows.length - 1
    const isTotal = Boolean(b.total_row) && last
    let cells = ''
    for (let c = 0; c < ncols; c++) {
      const v = c < r.length ? r[c] : ''
      const a = alignOf(align, c)
      let st: string
      if (isTotal) st = 'BwBTableTotal'
      else if (style === 'comparison' && c === 0) st = 'BwBTableTerm'
      else if (a === 'right') st = 'BwBTableNumber'
      else st = 'BwBTableText'
      const jc = a === 'center' || (a === 'right' && isTotal) ? a : null
      const keep = (together && !last) || (last && Boolean(b.caption))
      cells += doc.tc(ws[c] ?? 0, cellParas(doc, v, st, jc, keep), isTotal ? OFFWHITE : null, bord, tcmar)
    }
    trs.push('<w:tr><w:trPr><w:cantSplit/></w:trPr>' + cells + '</w:tr>')
  })
  const look =
    `<w:tblLook w:val="04A0" w:firstRow="1" w:lastRow="${b.total_row ? 1 : 0}" w:firstColumn="${style === 'comparison' ? 1 : 0}" ` +
    'w:lastColumn="0" w:noHBand="1" w:noVBand="1"/>'
  const desc = b.alt || b.caption ? `<w:tblCaption w:val="${esc(b.alt || b.caption || 'Table')}"/>` : ''
  const tbl =
    `<w:tbl><w:tblPr><w:tblStyle w:val="${tstyle}"/><w:tblW w:w="${TEXT_W}" w:type="dxa"/>` +
    `<w:tblLayout w:type="fixed"/>${look}${desc}</w:tblPr><w:tblGrid>` +
    ws.map((w) => `<w:gridCol w:w="${w}"/>`).join('') +
    '</w:tblGrid>' +
    trs.join('') +
    '</w:tbl>'
  doc.blockTable(tbl)
  if (b.caption) doc.add(caption(doc, 'Table', b.caption, b.source))
  else if (b.source) paragraph(doc, `Source: ${b.source}`, 'BwBSource')
  else doc.add(spacer(120))
}

/** Two-column label and value table (To, From, Date, Subject). */
export function keyValues(doc: Doc, rowsIn: KeyValueRows): void {
  const rows = Array.isArray(rowsIn) ? rowsIn : Object.entries(rowsIn)
  const ws = [2200, TEXT_W - 2200] as const
  const bord = `<w:tcBorders><w:bottom w:val="single" w:sz="4" w:space="0" w:color="${LIGHT}"/></w:tcBorders>`
  const mar =
    '<w:tcMar><w:top w:w="80" w:type="dxa"/><w:left w:w="0" w:type="dxa"/><w:bottom w:w="80" w:type="dxa"/>' +
    '<w:right w:w="120" w:type="dxa"/></w:tcMar>'
  let trs = ''
  for (const [k, v] of rows) {
    trs +=
      '<w:tr><w:trPr><w:cantSplit/></w:trPr>' +
      doc.tc(ws[0], para(doc.runs(String(k)), { style: 'BwBTableTerm' }), null, bord, mar) +
      doc.tc(ws[1], para(doc.runs(String(v)), { style: 'BwBTableText' }), null, bord, mar) +
      '</w:tr>'
  }
  doc.blockTable(doc.tblWrap('BwBPlain', ws, trs, '', '04A0', 'Document details'))
  doc.add(spacer(160))
}
