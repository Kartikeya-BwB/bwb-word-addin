// A tidy tree for org charts, built here rather than taken from a library. Each person's block is laid out from
// the bottom up: children side by side under their manager, who is centred over them. Long runs of people with no
// reports wrap into rows, and beyond eight they become a list in columns. The whole chart is then fitted to the
// content area, trying narrower boxes and tighter wrapping before it scales down.

import { CONTENT_AREA, GUTTER } from '../../brand/slideGrid.ts'
import type { Rect } from '../../geometry/index.ts'

export interface OrgPerson {
  name: string
  role: string
  children: OrgPerson[]
}

export interface OrgBox {
  person: OrgPerson
  /** 0 for the people at the top. */
  depth: number
  left: number
  top: number
  width: number
  height: number
}

/** Connector segments, each horizontal or vertical. */
export interface OrgLine {
  x1: number
  y1: number
  x2: number
  y2: number
}

export interface OrgLayout {
  boxes: OrgBox[]
  lines: OrgLine[]
  /** 1 unless the chart had to shrink to fit. */
  scale: number
}

interface Params {
  /** Box width and height. */
  w: number
  h: number
  /** Gap between neighbours, and between levels. */
  gx: number
  gy: number
  /** Most people with no reports in one row before they wrap. */
  rowMax: number
  /** Beyond this many people with no reports, a list in columns. */
  listAt: number
  /** People per list column. */
  perColumn: number
}

interface Block {
  width: number
  height: number
  /** The centre of the person's own box, from the block's left edge. */
  anchor: number
  boxes: Array<{ person: OrgPerson; depth: number; left: number; top: number }>
  lines: OrgLine[]
}

const LIST_INDENT = GUTTER
const LIST_GAP = GUTTER / 2

const hline = (x1: number, x2: number, y: number): OrgLine => ({ x1: Math.min(x1, x2), y1: y, x2: Math.max(x1, x2), y2: y })
const vline = (x: number, y1: number, y2: number): OrgLine => ({ x1: x, y1: Math.min(y1, y2), x2: x, y2: Math.max(y1, y2) })

function moved(b: Block, dx: number, dy: number): Pick<Block, 'boxes' | 'lines'> {
  return {
    boxes: b.boxes.map((x) => ({ ...x, left: x.left + dx, top: x.top + dy })),
    lines: b.lines.map((l) => ({ x1: l.x1 + dx, y1: l.y1 + dy, x2: l.x2 + dx, y2: l.y2 + dy })),
  }
}

function leaf(person: OrgPerson, depth: number, p: Params): Block {
  return { width: p.w, height: p.h, anchor: p.w / 2, boxes: [{ person, depth, left: 0, top: 0 }], lines: [] }
}

/** Children side by side, each with its own subtree, and the manager centred over the first and last of them. */
function row(person: OrgPerson, depth: number, p: Params): Block {
  const kids = person.children.map((c) => layoutPerson(c, depth + 1, p))
  const span = kids.reduce((s, k) => s + k.width, 0) + p.gx * (kids.length - 1)
  const inner = Math.max(p.w, span)
  const lefts: number[] = []
  let x = (inner - span) / 2
  for (const k of kids) {
    lefts.push(x)
    x += k.width + p.gx
  }
  const anchors = kids.map((k, i) => (lefts[i] ?? 0) + k.anchor)
  const centre = ((anchors[0] ?? inner / 2) + (anchors[anchors.length - 1] ?? inner / 2)) / 2
  // Keep the manager's box inside the block when the children lean to one side.
  const shift = Math.max(0, p.w / 2 - centre)
  const anchor = centre + shift
  const boxes: Block['boxes'] = [{ person, depth, left: anchor - p.w / 2, top: 0 }]
  const lines: OrgLine[] = []
  kids.forEach((k, i) => {
    const m = moved(k, (lefts[i] ?? 0) + shift, p.h + p.gy)
    boxes.push(...m.boxes)
    lines.push(...m.lines)
  })
  const drops = anchors.map((a) => a + shift)
  const bus = p.h + p.gy / 2
  lines.push(vline(anchor, p.h, bus))
  const lo = Math.min(anchor, ...drops)
  const hi = Math.max(anchor, ...drops)
  if (hi - lo > 0.01) lines.push(hline(lo, hi, bus))
  for (const a of drops) lines.push(vline(a, bus, p.h + p.gy))
  return {
    width: Math.max(inner + shift, anchor + p.w / 2), height: p.h + p.gy + Math.max(...kids.map((k) => k.height)), anchor, boxes, lines,
  }
}

/** People with no reports in balanced rows under the manager, joined by a spine that runs between boxes. */
function wrapped(person: OrgPerson, depth: number, p: Params): Block {
  const n = person.children.length
  const rows = Math.ceil(n / p.rowMax)
  const per = Math.ceil(n / rows)
  const rowWidth = (count: number): number => count * p.w + (count - 1) * p.gx
  const width = Math.max(p.w, rowWidth(per))
  const anchor = width / 2
  // With an odd number to a row, the middle of the row is a box: the spine runs in the gap to its right.
  const spine = per % 2 === 0 ? anchor : anchor + (p.w + p.gx) / 2
  const boxes: Block['boxes'] = [{ person, depth, left: anchor - p.w / 2, top: 0 }]
  const lines: OrgLine[] = [vline(anchor, p.h, p.h + p.gy / 2)]
  for (let r = 0; r < rows; r += 1) {
    const members = person.children.slice(r * per, (r + 1) * per)
    const top = p.h + p.gy + r * (p.h + p.gy)
    const bus = top - p.gy / 2
    const left0 = (width - rowWidth(members.length)) / 2
    const centres = members.map((_, i) => left0 + i * (p.w + p.gx) + p.w / 2)
    members.forEach((m, i) => boxes.push({ person: m, depth: depth + 1, left: (centres[i] ?? 0) - p.w / 2, top }))
    lines.push(hline(Math.min(spine, anchor, ...centres), Math.max(spine, anchor, ...centres), bus))
    for (const c of centres) lines.push(vline(c, bus, top))
  }
  if (rows > 1) lines.push(vline(spine, p.h + p.gy / 2, p.h + p.gy + (rows - 1) * (p.h + p.gy) - p.gy / 2))
  return { width, height: p.h + rows * (p.h + p.gy), anchor, boxes, lines }
}

/** Many people with no reports as a list in columns, each column with its own spine and a tick to each person. */
function list(person: OrgPerson, depth: number, p: Params): Block {
  const n = person.children.length
  const cols = Math.ceil(n / Math.max(1, p.perColumn))
  const per = Math.ceil(n / cols)
  const colWidth = LIST_INDENT + p.w
  const inner = cols * colWidth + (cols - 1) * p.gx
  const width = Math.max(p.w, inner)
  const anchor = width / 2
  const bus = p.h + p.gy / 2
  const boxes: Block['boxes'] = [{ person, depth, left: anchor - p.w / 2, top: 0 }]
  const lines: OrgLine[] = [vline(anchor, p.h, bus)]
  const spines: number[] = []
  for (let c = 0; c < cols; c += 1) {
    const members = person.children.slice(c * per, (c + 1) * per)
    const x0 = (width - inner) / 2 + c * (colWidth + p.gx)
    const spine = x0 + LIST_INDENT / 2
    spines.push(spine)
    members.forEach((m, i) => {
      const top = p.h + p.gy + i * (p.h + LIST_GAP)
      boxes.push({ person: m, depth: depth + 1, left: x0 + LIST_INDENT, top })
      lines.push(hline(spine, x0 + LIST_INDENT, top + p.h / 2))
    })
    lines.push(vline(spine, bus, p.h + p.gy + (members.length - 1) * (p.h + LIST_GAP) + p.h / 2))
  }
  lines.push(hline(Math.min(anchor, ...spines), Math.max(anchor, ...spines), bus))
  return { width, height: p.h + p.gy + per * p.h + (per - 1) * LIST_GAP, anchor, boxes, lines }
}

function layoutPerson(person: OrgPerson, depth: number, p: Params): Block {
  const n = person.children.length
  if (!n) return leaf(person, depth, p)
  const onlyLeaves = person.children.every((c) => !c.children.length)
  if (onlyLeaves && n > p.listAt) return list(person, depth, p)
  if (onlyLeaves && n > p.rowMax) return wrapped(person, depth, p)
  return row(person, depth, p)
}

/** Several people at the top sit side by side, twice the usual gap apart. */
function forest(roots: readonly OrgPerson[], p: Params): Block {
  const blocks = roots.map((r) => layoutPerson(r, 0, p))
  const out: Block = { width: 0, height: 0, anchor: 0, boxes: [], lines: [] }
  let x = 0
  for (const b of blocks) {
    const m = moved(b, x, 0)
    out.boxes.push(...m.boxes)
    out.lines.push(...m.lines)
    x += b.width + 2 * p.gx
  }
  out.width = Math.max(0, x - 2 * p.gx)
  out.height = Math.max(0, ...blocks.map((b) => b.height))
  return out
}

/** Box widths and wrapping to try, roomiest first. */
function candidates(area: Rect): Params[] {
  const h = 40
  const gy = GUTTER * 1.5
  const perColumn = Math.max(1, Math.floor((area.height - h - gy + LIST_GAP) / (h + LIST_GAP)))
  const out: Params[] = []
  for (const w of [132, 116, 100, 88]) {
    for (const rowMax of [6, 4, 3]) out.push({ w, h, gx: GUTTER, gy, rowMax, listAt: 8, perColumn })
  }
  return out
}

/** Lays out the chart and fits it to the area: centred across, from the top, scaled down only as a last resort. */
export function layoutOrgChart(roots: readonly OrgPerson[], area: Rect = CONTENT_AREA): OrgLayout {
  const tries = candidates(area)
  let chosen: { block: Block; p: Params } | null = null
  for (const p of tries) {
    const block = forest(roots, p)
    if (block.width <= area.width + 0.01 && block.height <= area.height + 0.01) {
      chosen = { block, p }
      break
    }
  }
  const p = chosen?.p ?? tries[tries.length - 1] ?? { w: 88, h: 40, gx: GUTTER, gy: GUTTER * 1.5, rowMax: 3, listAt: 8, perColumn: 6 }
  const block = chosen?.block ?? forest(roots, p)
  const scale = chosen ? 1 : Math.min(1, area.width / Math.max(1, block.width), area.height / Math.max(1, block.height))
  const ox = area.left + (area.width - block.width * scale) / 2
  const oy = area.top
  return {
    scale,
    boxes: block.boxes.map((b) => ({ person: b.person, depth: b.depth, left: ox + b.left * scale, top: oy + b.top * scale, width: p.w * scale, height: p.h * scale })),
    lines: block.lines.map((l) => ({ x1: ox + l.x1 * scale, y1: oy + l.y1 * scale, x2: ox + l.x2 * scale, y2: oy + l.y2 * scale })),
  }
}
