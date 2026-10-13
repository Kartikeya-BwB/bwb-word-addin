// The pane's tabs for each app. Word and PowerPoint share one pane: each opens on its own set, and every tab's code
// loads only when the tab is first opened, so the pane starts quickly in both.

import { lazy } from 'react'
import type { ComponentType, LazyExoticComponent } from 'react'
import { FilePlus2, LayoutPanelTop, Palette, PenLine, Presentation, ShieldCheck, SquarePlus, Wrench } from 'lucide-react'
import type { LucideIcon } from 'lucide-react'
import type { Host } from '../host.ts'

export interface TabDef {
  id: string
  label: string
  Icon: LucideIcon
  Panel: LazyExoticComponent<ComponentType>
}

/** React.lazy for a named export. */
function load<M>(loader: () => Promise<M>, pick: (m: M) => ComponentType): LazyExoticComponent<ComponentType> {
  return lazy(async () => ({ default: pick(await loader()) }))
}

const WORD_TABS: TabDef[] = [
  { id: 'start', label: 'Start', Icon: FilePlus2, Panel: load(() => import('./tabs/StartTab.tsx'), (m) => m.StartTab) },
  { id: 'write', label: 'Write', Icon: PenLine, Panel: load(() => import('./tabs/WriteTab.tsx'), (m) => m.WriteTab) },
  { id: 'insert', label: 'Insert', Icon: SquarePlus, Panel: load(() => import('./tabs/InsertTab.tsx'), (m) => m.InsertTab) },
  { id: 'tools', label: 'Tools', Icon: Wrench, Panel: load(() => import('./tabs/ToolsTab.tsx'), (m) => m.ToolsTab) },
  { id: 'check', label: 'Check', Icon: ShieldCheck, Panel: load(() => import('./tabs/CheckTab.tsx'), (m) => m.CheckTab) },
  { id: 'brand', label: 'Brand', Icon: Palette, Panel: load(() => import('./tabs/BrandTab.tsx'), (m) => m.BrandTab) },
]

const POWERPOINT_TABS: TabDef[] = [
  { id: 'slides', label: 'Slides', Icon: Presentation, Panel: load(() => import('./tabs/SlidesPanel.tsx'), (m) => m.SlidesPanel) },
  { id: 'arrange', label: 'Arrange', Icon: LayoutPanelTop, Panel: load(() => import('./ppt/ArrangeTab.tsx'), (m) => m.ArrangeTab) },
  { id: 'insert', label: 'Insert', Icon: SquarePlus, Panel: load(() => import('./tabs/SlideInsertTab.tsx'), (m) => m.SlideInsertTab) },
  { id: 'check', label: 'Check', Icon: ShieldCheck, Panel: load(() => import('./ppt/DeckCheckTab.tsx'), (m) => m.DeckCheckTab) },
  { id: 'brand', label: 'Brand', Icon: Palette, Panel: load(() => import('./tabs/BrandTab.tsx'), (m) => m.BrandTab) },
]

export function tabsFor(host: Host): TabDef[] {
  return host === 'powerpoint' ? POWERPOINT_TABS : WORD_TABS
}

/** What the masthead says the pane is for. */
export const HOST_TAGLINE: Record<Host, string> = {
  word: 'Word documents in the BwB format',
  powerpoint: 'Slides in the BwB format',
}
