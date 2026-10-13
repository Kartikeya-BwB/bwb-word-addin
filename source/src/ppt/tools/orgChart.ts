// An org chart from a typed or pasted list: "Name, Role" lines indented under their manager, or
// "Name | Role | Reports to" lines. Laid out by tools/orgLayout.ts and drawn as rounded boxes and connector lines.

import { SLIDE_TYPE } from '../../brand/slideGrid.ts'
import { COLOURS } from '../../brand/tokens.ts'
import { UserFacingError } from '../../word/run.ts'
import type { TextSpec } from '../specs.ts'
import type { DrawShape } from './draw.ts'
import { insertDrawing, segment } from './draw.ts'
import type { OrgLayout, OrgPerson } from './orgLayout.ts'
import { layoutOrgChart } from './orgLayout.ts'

export type { OrgPerson } from './orgLayout.ts'

/** More people than this cannot be read on one slide. */
export const MAX_PEOPLE = 40

const BULLET = /^(?:[-*•–]\s+)/

function person(text: string, separator: RegExp): OrgPerson {
  const cut = separator.exec(text)
  const name = (cut ? text.slice(0, cut.index) : text).trim()
  const role = cut ? text.slice(cut.index + cut[0].length).trim() : ''
  return { name, role, children: [] }
}

/** Indented lines: each line reports to the nearest line above it that is indented less. Tabs count four spaces. */
function indented(lines: readonly string[]): OrgPerson[] {
  const roots: OrgPerson[] = []
  const stack: Array<{ indent: number; person: OrgPerson }> = []
  for (const raw of lines) {
    const expanded = raw.replace(/\t/g, '    ')
    const indent = expanded.length - expanded.trimStart().length
    const p = person(expanded.trim().replace(BULLET, ''), /\s*,\s*/)
    while (stack.length && (stack[stack.length - 1]?.indent ?? 0) >= indent) stack.pop()
    const parent = stack[stack.length - 1]?.person
    if (parent) parent.children.push(p)
    else roots.push(p)
    stack.push({ indent, person: p })
  }
  return roots
}

/** "Name | Role | Reports to" lines, in any order. A first line of column headings is skipped. */
function piped(lines: readonly string[]): OrgPerson[] {
  const rows = lines.map((l) => l.split('|').map((c) => c.trim()))
  if (/^name$/i.test(rows[0]?.[0] ?? '')) rows.shift()
  const people = rows.map((r, i) => {
    const name = r[0] ?? ''
    if (!name) throw new UserFacingError(`Line ${i + 1} has no name. Write each line as Name | Role | Reports to.`)
    return { person: { name, role: r[1] ?? '', children: [] } as OrgPerson, manager: r[2] ?? '' }
  })
  const byName = new Map<string, OrgPerson>()
  for (const { person: p } of people) {
    const key = p.name.toLowerCase()
    if (byName.has(key)) throw new UserFacingError(`“${p.name}” is on the list twice. Give each person a name of their own.`)
    byName.set(key, p)
  }
  const managerOf = new Map<OrgPerson, OrgPerson>()
  const roots: OrgPerson[] = []
  for (const { person: p, manager } of people) {
    if (!manager) {
      roots.push(p)
      continue
    }
    const boss = byName.get(manager.toLowerCase())
    if (!boss) throw new UserFacingError(`“${p.name}” reports to “${manager}”, who is not on the list.`)
    if (boss === p) throw new UserFacingError(`“${p.name}” cannot report to themselves.`)
    boss.children.push(p)
    managerOf.set(p, boss)
  }
  for (const { person: p } of people) {
    const seen = new Set<OrgPerson>()
    for (let at: OrgPerson | undefined = p; at; at = managerOf.get(at)) {
      if (seen.has(at)) throw new UserFacingError(`The reporting lines go round in a circle at “${at.name}”.`)
      seen.add(at)
    }
  }
  return roots
}

const count = (people: readonly OrgPerson[]): number => people.reduce((n, p) => n + 1 + count(p.children), 0)

/** Reads the list. Pipes on any line mean the "Reports to" form; otherwise indentation gives the reporting lines. */
export function parseOrgChart(text: string): OrgPerson[] {
  const lines = text.split(/\r\n|[\r\n]/).filter((l) => l.trim())
  if (!lines.length) throw new UserFacingError('Type or paste the people first, one to a line.')
  const roots = lines.some((l) => l.includes('|')) ? piped(lines) : indented(lines)
  const total = count(roots)
  if (total > MAX_PEOPLE) throw new UserFacingError(`${total} people will not fit on one slide. Split the chart over several slides, up to ${MAX_PEOPLE} a slide.`)
  return roots
}

const APTOS = SLIDE_TYPE.body.font

/** Boxes and lines as shapes: lines at the back, the people at the top in navy, everyone else white with a navy edge. */
export function orgChartShapes(layout: OrgLayout): DrawShape[] {
  const size = Math.max(SLIDE_TYPE.minAny, Math.round(SLIDE_TYPE.small.size * layout.scale * 2) / 2)
  const pad = 4 * layout.scale
  const lines = layout.lines.map((l, i) => segment(`BwB org chart line ${i + 1}`, l.x1, l.y1, l.x2, l.y2, { colour: COLOURS.charcoal, weight: 1 }))
  const boxes = layout.boxes.map((b): DrawShape => {
    const top = b.depth === 0
    const paragraphs: TextSpec['paragraphs'] = [{ text: b.person.name, font: { name: APTOS, size, colour: top ? COLOURS.white : COLOURS.navy, bold: true } }]
    if (b.person.role) paragraphs.push({ text: b.person.role, font: { name: APTOS, size, colour: top ? COLOURS.white : COLOURS.charcoal, bold: false } })
    return {
      kind: 'geometric', geometry: 'RoundRectangle', name: `BwB org chart: ${b.person.name}`, left: b.left, top: b.top, width: b.width, height: b.height,
      fill: { colour: top ? COLOURS.navy : COLOURS.white }, line: top ? null : { colour: COLOURS.navy, weight: 1 },
      text: { paragraphs, align: 'Center', vertical: 'Middle', margins: { left: pad, right: pad, top: pad / 2, bottom: pad / 2 }, wrap: true },
    }
  })
  return [...lines, ...boxes]
}

/** Draws the chart in the content area of the current slide, grouped and tagged BWB_ELEMENT=org-chart. */
export async function insertOrgChart(text: string): Promise<string> {
  const roots = parseOrgChart(text)
  const layout = layoutOrgChart(roots)
  await insertDrawing('org-chart', 'BwB org chart', orgChartShapes(layout), 'The org chart')
  const people = count(roots)
  const fit = layout.scale < 1 ? ` It was scaled to fit, so the text is ${Math.max(SLIDE_TYPE.minAny, Math.round(SLIDE_TYPE.small.size * layout.scale * 2) / 2)} pt.` : ''
  return `Org chart of ${people} ${people === 1 ? 'person' : 'people'} added.${fit}`
}
