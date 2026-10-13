// Every tool on the ribbon of the current app, as one searchable list for the pane's "Find a tool" box. Large
// toolbars hide their tools in menus; searching by what you want to do finds them again.

import wordRibbon from './ribbon-word.json'
import pptRibbon from './ribbon-powerpoint.json'
import type { Host } from '../host.ts'

export interface Tool {
  id: string
  label: string
  /** Where it lives on the ribbon: "Insert, Blocks". */
  path: string
  tip: string
  fn?: string
  pane?: string
}

interface RibbonEntry { id: string; label: string; tip: string; fn?: string; pane?: string; items?: RibbonEntry[] }
interface Ribbon { groups: Array<{ id: string; label: string; controls: RibbonEntry[] }> }

const RIBBONS: Record<Host, Ribbon> = { word: wordRibbon as Ribbon, powerpoint: pptRibbon as Ribbon }

export function toolsFor(host: Host): Tool[] {
  const out: Tool[] = []
  const add = (e: RibbonEntry, path: string): void => {
    if (!e.fn && !e.pane) return
    out.push({ id: e.id, label: e.label, path, tip: e.tip, ...(e.fn ? { fn: e.fn } : {}), ...(e.pane ? { pane: e.pane } : {}) })
  }
  for (const group of RIBBONS[host].groups) {
    for (const control of group.controls) {
      add(control, group.label)
      for (const item of control.items ?? []) add(item, `${group.label}, ${control.label}`)
    }
  }
  return out
}

/** Ranks tools for a query: whole words in the label count most, then the menu, then the description. */
export function searchTools(tools: Tool[], query: string, limit = 8): Tool[] {
  const words = query.toLowerCase().split(/\s+/).filter(Boolean)
  if (!words.length) return []
  const scored = tools.map((tool) => {
    const label = tool.label.toLowerCase()
    const path = tool.path.toLowerCase()
    const tip = tool.tip.toLowerCase()
    let score = 0
    for (const w of words) {
      if (new RegExp(`\\b${w.replace(/[.*+?^${}()|[\]\\]/g, '\\$&')}`).test(label)) score += 6
      else if (label.includes(w)) score += 4
      else if (path.includes(w)) score += 2
      else if (tip.includes(w)) score += 1
      else return { tool, score: 0 }
    }
    return { tool, score }
  })
  return scored.filter((s) => s.score > 0).sort((a, b) => b.score - a.score || a.tool.label.localeCompare(b.tool.label)).slice(0, limit).map((s) => s.tool)
}
