import {
  BookCheck, BookOpen, ChartGantt, ClipboardList, File, FileMinus, FileText, Flag, Landmark, Lightbulb, Mail,
  Microscope, Newspaper, Route, Scale, ScanSearch, Send,
} from 'lucide-react'
import type { LucideIcon } from 'lucide-react'

// One icon per document type, named in src/brand/docTypes.ts; imported one by one to keep the bundle small.
const ICONS: Record<string, LucideIcon> = {
  'inception-report': Flag,
  'final-report': BookCheck,
  'research-report': Microscope,
  'gap-analysis': ScanSearch,
  'progress-report': ChartGantt,
  roadmap: Route,
  eoi: Send,
  'guidance-note': BookOpen,
  'technical-proposal': ClipboardList,
  'feasibility-study': Scale,
  'policy-brief': Landmark,
  'briefing-note': FileText,
  'concept-note': Lightbulb,
  'white-paper': Newspaper,
  memo: Mail,
  'blank-report': File,
  'blank-memo': FileMinus,
}

export function DocTypeIcon({ id, size = 18 }: { id: string; size?: number }) {
  const Icon = ICONS[id] ?? FileText
  return <Icon size={size} strokeWidth={1.8} color="var(--navy)" aria-hidden />
}
