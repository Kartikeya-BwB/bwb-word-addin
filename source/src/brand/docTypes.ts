// The document types the add-in can start, chosen from the BwB Deliverables Hub catalogue
// (document-types research, 10 October 2026) in order of how often BwB produces them.

import { SKELETONS } from './skeletons.ts'
import type { Block } from '../docx/types.ts'

export type Layout = 'report' | 'memo'

export interface DocType {
  id: string
  name: string
  family: string
  icon: string
  summary: string
  length?: string
  layout: Layout
  /** Client-facing documents end with the BwB disclaimer; bids and internal notes do not. */
  disclaimer: boolean
  checks: string[]
  skeleton: () => Block[]
}

function skeleton(id: string): () => Block[] {
  const make = SKELETONS[id]
  if (!make) throw new Error(`No skeleton for document type ${id}`)
  return make
}

export const DOC_TYPES: DocType[] = [
  {
    id: 'inception-report', name: 'Inception report', family: 'Inception and planning', icon: 'flag', layout: 'report', disclaimer: true,
    summary: 'Restates the ToR, sets out the approach, the work plan and the deliverable schedule',
    length: '5 to 15 pages where the ToR caps it',
    checks: [
      'Objectives and scope restated against the ToR, including clarifications agreed at kick-off',
      'Methodology given for each deliverable',
      'Work plan and Gantt dates match the contract',
      'Team and roles; stakeholder map; data plan and data requests; risks and mitigation',
      'Deliverable schedule matches contract dates and payment shares',
    ],
    skeleton: skeleton('inception-report'),
  },
  {
    id: 'final-report', name: 'Final report', family: 'Contract and progress reporting', icon: 'book-check', layout: 'report', disclaimer: true,
    summary: 'Consolidates earlier deliverables, with outcomes against objectives and recommendations',
    checks: [
      'Consolidates earlier deliverables, with outcomes against objectives',
      'Recommendations are actionable',
      'Annexes are complete',
      'A comments log shows how the client’s feedback on the draft was addressed',
    ],
    skeleton: skeleton('final-report'),
  },
  {
    id: 'research-report', name: 'Research or analytical report', family: 'Diagnostic and analytical', icon: 'microscope', layout: 'report', disclaimer: true,
    summary: 'Answers a research question with a stated method and current sources',
    length: '40 to 70 pages unless the ToR caps it',
    checks: ['The research question and method are stated', 'Sources are current and claims are proportionate', 'There is a limits section'],
    skeleton: skeleton('research-report'),
  },
  {
    id: 'gap-analysis', name: 'Diagnostic or gap analysis', family: 'Diagnostic and analytical', icon: 'scan-search', layout: 'report', disclaimer: true,
    summary: 'Reviews policy, benchmarks peers and prioritises gaps, with a recommendation for each',
    checks: [
      'The policy and regulatory review is current, with dated citations',
      'Peer benchmarking is included',
      'Interview and survey numbers are counted',
      'Gaps are prioritised, and every recommendation maps to a gap',
      'Country data is consistent across sections',
    ],
    skeleton: skeleton('gap-analysis'),
  },
  {
    id: 'progress-report', name: 'Progress or interim report', family: 'Contract and progress reporting', icon: 'chart-gantt', layout: 'report', disclaimer: true,
    summary: 'Progress against plan for a reporting period, with deviations explained',
    checks: [
      'The reporting period is correct',
      'Activities are compared against plan, and deviations are explained',
      'KPIs and a financial section where required',
      'The funder’s template is used',
    ],
    skeleton: skeleton('progress-report'),
  },
  {
    id: 'roadmap', name: 'Roadmap, action plan or strategy', family: 'Strategy, frameworks and guidance', icon: 'route', layout: 'report', disclaimer: true,
    summary: 'A vision, pillars and a costed, owned and sequenced action plan',
    checks: [
      'Built on a diagnostic and peer benchmarking',
      'Consultations held',
      'Every action has an owner, timing, sequencing and dependencies',
      'An M&E framework, and validation planned',
    ],
    skeleton: skeleton('roadmap'),
  },
  {
    id: 'eoi', name: 'Expression of interest', family: 'Bids and proposals', icon: 'send', layout: 'memo', disclaimer: false,
    summary: 'The four-section BwB EOI: qualifications, capabilities, core business, key staff',
    length: 'Shortlisted BwB EOIs run 5,400 to 8,400 words',
    checks: [
      'Every criterion is answered under its own label',
      'The local partner is named, with its task',
      'No borrowed boilerplate and no fee values',
      'Attachments meet size limits',
    ],
    skeleton: skeleton('eoi'),
  },
  {
    id: 'guidance-note', name: 'Guidance note or guide', family: 'Strategy, frameworks and guidance', icon: 'book-open', layout: 'report', disclaimer: true,
    summary: 'Steps in order, with worked examples and a glossary, for a named user',
    checks: ['Usable by the named user', 'Steps in order, with worked examples and a glossary', 'Delivered as an editable file where the client owns it'],
    skeleton: skeleton('guidance-note'),
  },
  {
    id: 'technical-proposal', name: 'Technical proposal', family: 'Bids and proposals', icon: 'clipboard-list', layout: 'report', disclaimer: false,
    summary: 'Understanding, method, work plan and team, mapped to a compliance matrix',
    length: 'Ten pages is the commonest limit',
    checks: [
      'The compliance matrix maps every requirement and criterion',
      'Page, word and format limits are met',
      'Team, dates and person-days are consistent across every form',
      'No price anywhere in the technical file when the funder uses two envelopes',
    ],
    skeleton: skeleton('technical-proposal'),
  },
  {
    id: 'feasibility-study', name: 'Feasibility study', family: 'Diagnostic and analytical', icon: 'scale', layout: 'report', disclaimer: true,
    summary: 'Options appraised against stated criteria, ending in one recommendation',
    checks: [
      'Options are appraised against stated criteria',
      'Financial viability is backed by the model, with matching figures',
      'Legal, regulatory and institutional feasibility is covered, along with risks',
      'It ends in one clear recommendation',
    ],
    skeleton: skeleton('feasibility-study'),
  },
  {
    id: 'policy-brief', name: 'Policy brief', family: 'Strategy, frameworks and guidance', icon: 'landmark', layout: 'memo', disclaimer: true,
    summary: 'Key messages and directive recommendations for named institutions',
    length: '1 to 8 pages',
    checks: ['Recommendations are directive and addressed to named institutions', 'Politically neutral wording; within mandate', 'The length cap is met'],
    skeleton: skeleton('policy-brief'),
  },
  {
    id: 'briefing-note', name: 'Briefing note or one-pager', family: 'Strategy, frameworks and guidance', icon: 'file-text', layout: 'memo', disclaimer: false,
    summary: 'One question, one answer and the decision requested',
    checks: ['One question, one answer', 'The decision requested is explicit', 'The length limit is met'],
    skeleton: skeleton('briefing-note'),
  },
  {
    id: 'concept-note', name: 'Concept note or grant proposal', family: 'Bids and proposals', icon: 'lightbulb', layout: 'memo', disclaimer: false,
    summary: 'The problem, the approach and the theory of change for a funder',
    checks: ['The funder’s template and word limits are followed', 'The theory of change is coherent'],
    skeleton: skeleton('concept-note'),
  },
  {
    id: 'white-paper', name: 'White paper or thought piece', family: 'Knowledge and thought leadership', icon: 'newspaper', layout: 'report', disclaimer: true,
    summary: 'A clear thesis backed by evidence, for publication',
    checks: ['A clear thesis, backed by evidence and sources', 'Copyright notice included', 'AI-detection scan recorded', 'Co-author sign-offs and partner logo consent'],
    skeleton: skeleton('white-paper'),
  },
  {
    id: 'memo', name: 'Internal memo', family: 'Internal', icon: 'mail', layout: 'memo', disclaimer: false,
    summary: 'A short internal note with a details table',
    checks: ['The recommendation or decision is stated'],
    skeleton: skeleton('memo'),
  },
  {
    id: 'blank-report', name: 'Blank BwB report', family: 'Blank', icon: 'file', layout: 'report', disclaimer: true,
    summary: 'Cover, contents, numbered headings and the disclaimer, with no prompts',
    checks: [],
    skeleton: skeleton('blank-report'),
  },
  {
    id: 'blank-memo', name: 'Blank BwB memo', family: 'Blank', icon: 'file-minus', layout: 'memo', disclaimer: false,
    summary: 'A title block and details table, with no prompts',
    checks: [],
    skeleton: skeleton('blank-memo'),
  },
]

export function findDocType(id: string | undefined): DocType | undefined {
  return id ? DOC_TYPES.find((t) => t.id === id) : undefined
}

export function searchDocTypes(query: string): DocType[] {
  const q = query.trim().toLowerCase()
  if (!q) return DOC_TYPES
  return DOC_TYPES.filter((t) => `${t.name} ${t.family} ${t.summary}`.toLowerCase().includes(q))
}
