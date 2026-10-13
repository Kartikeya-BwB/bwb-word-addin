// Starting structures for new documents, written in the BwB content model (the same model the BwB Document
// Format skill builds from). Guidance blocks become tagged content controls that the Pre-QA check reports
// until the author replaces them; {field:Label} tokens become placeholders the author fills in.

import type { Block } from '../docx/types.ts'

const h1 = (text: string): Block => ({ type: 'heading', level: 1, text })
const h2 = (text: string): Block => ({ type: 'heading', level: 2, text })
const g = (text: string): Block => ({ type: 'guidance', text })
const actions = (items: string[]): Block => ({ type: 'bullets', style: 'arrow', items })

const riskTable: Block = {
  type: 'table',
  header: ['Risk', 'Likelihood', 'Impact', 'Mitigation', 'Owner'],
  rows: [['{field:Risk}', '{field:Low, medium or high}', '{field:Low, medium or high}', '{field:Mitigation}', '{field:Owner}']],
  caption: 'Risks and mitigation',
}

const executiveSummary = (guidance: string): Block[] => [h1('Executive summary'), g(guidance)]

export const SKELETONS: Record<string, () => Block[]> = {
  'inception-report': () => [
    ...executiveSummary('Summarise the assignment, the agreed approach and the first deliverables in half a page. The executive summary only summarises.'),
    h1('Introduction'),
    g('Background to the assignment, the client and the purpose of this report.'),
    h1('Objectives and scope'),
    g('Restate the objectives and scope against the ToR, including the clarifications agreed at kick-off. Say what is out of scope.'),
    h1('Approach and methodology'),
    g('Give the methodology for each deliverable: sources, analytical framework, consultations and limits.'),
    h1('Work plan and timeline'),
    g('Activities and a Gantt chart. Every date must match the contract.'),
    {
      type: 'table',
      header: ['Deliverable', 'Description', 'Due date', 'Payment share'],
      rows: [['{field:D1}', '{field:Description}', '{field:Date}', '{field:Per cent}']],
      caption: 'Deliverable schedule',
    },
    h1('Team and governance'),
    g('Team members and roles, reporting lines and the client focal point.'),
    h1('Stakeholders and data'),
    g('Stakeholder map, data plan and the data requests to the client.'),
    h1('Risks and mitigation'),
    riskTable,
    h1('Next steps'),
    actions(['{field:Next step}']),
    h1('Annexes'),
    g('If the ToR asks for it, add an annotated table of contents for the final report, a definition of success or a theory of change.'),
  ],

  'final-report': () => [
    ...executiveSummary('The findings and recommendations in one or two pages, for a reader who reads nothing else.'),
    h1('Introduction'),
    g('Objectives, scope and the period covered. Link to the earlier deliverables this report consolidates.'),
    h1('Approach and methodology'),
    g('How the work was done, the sources and the limitations.'),
    h1('Findings'),
    g('One sub-section per finding or theme. Compare outcomes against the objectives.'),
    h2('{field:Finding}'),
    h1('Recommendations'),
    g('Make every recommendation actionable, with an owner and timing.'),
    { type: 'table', header: ['Recommendation', 'Owner', 'Timing'], rows: [['{field:Recommendation}', '{field:Owner}', '{field:Timing}']], caption: 'Recommendations' },
    h1('Next steps'),
    actions(['{field:Next step}']),
    h1('Annexes'),
    g('Include a comments log (comment, response, location) when this version answers client comments on a draft.'),
  ],

  'research-report': () => [
    ...executiveSummary('The question, the answer and what it means, in one page.'),
    h1('Introduction'),
    g('State the research question and why it matters now.'),
    h1('Method and data'),
    g('Sources, analytical approach and data limits. Date every data source.'),
    h1('Findings'),
    h2('{field:Theme}'),
    g('Keep claims proportionate to the evidence and cite current sources.'),
    h1('Conclusions and recommendations'),
    h1('Limitations'),
    g('What the analysis cannot show and why.'),
    h1('References'),
  ],

  'gap-analysis': () => [
    ...executiveSummary('The most important gaps and what to do about them.'),
    h1('Introduction'),
    h1('Policy and regulatory review'),
    g('Current policy and regulation, with dated citations.'),
    h1('Current state'),
    g('State interview and survey numbers (for example, 24 interviews and 112 survey responses).'),
    h1('Peer benchmarking'),
    h1('Gaps'),
    { type: 'table', header: ['Gap', 'Evidence', 'Priority'], rows: [['{field:Gap}', '{field:Evidence}', '{field:High, medium or low}']], caption: 'Gaps by priority' },
    h1('Recommendations'),
    g('Every recommendation maps to a gap.'),
    { type: 'table', header: ['Recommendation', 'Gap addressed', 'Owner', 'Timing'], rows: [['{field:Recommendation}', '{field:Gap}', '{field:Owner}', '{field:Timing}']], caption: 'Recommendations and the gaps they address' },
    h1('Annexes'),
  ],

  'feasibility-study': () => [
    ...executiveSummary('The recommendation first, then the reasons.'),
    h1('Introduction and objectives'),
    h1('Context and market'),
    h1('Options considered'),
    g('State the appraisal criteria before the options.'),
    h1('Options appraisal'),
    { type: 'table', style: 'comparison', header: ['Criterion', 'Option A', 'Option B', 'Option C'], rows: [['{field:Criterion}', '', '', '']], caption: 'Options appraisal' },
    h1('Financial viability'),
    g('Every figure here must match the financial model.'),
    h1('Legal, regulatory and institutional feasibility'),
    h1('Risks'),
    riskTable,
    h1('Recommendation'),
    g('End with one clear recommendation.'),
    h1('Annexes'),
  ],

  roadmap: () => [
    ...executiveSummary('The vision, the pillars and the first actions.'),
    h1('Introduction'),
    h1('Diagnostic and benchmarking'),
    g('Summarise the diagnostic and the peer benchmarking the roadmap builds on.'),
    h1('Vision and pillars'),
    h1('Action plan'),
    g('Every action has an owner, timing, sequencing, dependencies and a cost. Add a RACI where the client asks for one.'),
    { type: 'table', header: ['Action', 'Owner', 'Timing', 'Dependencies', 'Cost'], rows: [['{field:Action}', '{field:Owner}', '{field:Timing}', '{field:Dependencies}', '{field:Cost}']], caption: 'Action plan' },
    h1('Monitoring and evaluation'),
    h1('Validation'),
    g('How and when the roadmap will be validated with stakeholders.'),
    h1('Annexes'),
  ],

  'guidance-note': () => [
    h1('About this guide'),
    g('Who the guide is for and what it helps them do.'),
    h1('Introduction'),
    h1('Key concepts'),
    h1('Step by step'),
    { type: 'numbered', items: ['{field:Step}'] },
    h1('Worked example'),
    h1('Further resources'),
    h1('Glossary'),
  ],

  'policy-brief': () => [
    { type: 'callout', title: 'Key messages', text: '', bullets: ['{field:Key message}'] },
    h1('The issue'),
    h1('Evidence'),
    h1('Policy options'),
    h1('Recommendations'),
    g('Make recommendations directive and address each to a named institution. Keep wording politically neutral and within the mandate.'),
    h1('References'),
  ],

  'briefing-note': () => [
    h1('Purpose'),
    g('One question, stated plainly.'),
    h1('Summary'),
    g('The answer in three sentences.'),
    h1('Background'),
    h1('Analysis'),
    { type: 'callout', title: 'Decision requested', text: '{field:The decision you need}' },
    h1('Next steps'),
    actions(['{field:Next step}']),
  ],

  'progress-report': () => [
    h1('Summary'),
    h1('Reporting period and scope'),
    g('State the reporting period exactly as the contract or grant defines it.'),
    h1('Progress against plan'),
    { type: 'table', header: ['Task or milestone', 'Planned', 'Status', 'Comment'], rows: [['{field:Task}', '{field:Date}', '{field:Status}', '{field:Comment}']], caption: 'Progress against plan' },
    h1('Deviations and corrective actions'),
    g('Explain every deviation from the plan.'),
    h1('Key performance indicators'),
    g('Where the funder requires them: definitions, baselines and this period’s values.'),
    h1('Risks and issues'),
    h1('Plan for the next period'),
  ],

  eoi: () => [
    g('Reproduce each section prompt from the procurement notice in bold before the section. Shortlisted BwB EOIs run 5,400 to 8,400 words. Never cite fee values.'),
    h1('Qualifications in the field'),
    g('Open with “Bankers without Boundaries (BwB) is pleased to submit this Expression of Interest for the” and the exact assignment title. Sub-section headers mirror the evaluation criteria. End with the MDB and Global Network paragraphs.'),
    h1('Technical and managerial capabilities'),
    h1('Core business and years in operation'),
    h1('Qualifications of key staff'),
    g('Two paragraphs per person. Label the team lead.'),
  ],

  'technical-proposal': () => [
    g('Follow the funder’s form numbering and limits (for example TECH-4 or Form E). Never put a price in the technical file when the funder uses two envelopes.'),
    h1('Understanding of the assignment'),
    h1('Approach and methodology'),
    h1('Work plan and schedule'),
    h1('Team and organisation'),
    g('Team, dates and person-days must match every other form.'),
    h1('Relevant experience'),
    h1('Quality assurance and risk management'),
    h1('Annex: compliance matrix'),
    { type: 'table', header: ['Requirement', 'Where it is addressed'], rows: [['{field:Requirement}', '{field:Section}']], caption: 'Compliance matrix' },
  ],

  'concept-note': () => [
    h1('Summary'),
    h1('Problem and opportunity'),
    h1('Proposed approach'),
    g('Set out the theory of change.'),
    h1('Activities and outputs'),
    h1('Partners and roles'),
    h1('Budget summary'),
    h1('Expected impact'),
    h1('Next steps'),
  ],

  'white-paper': () => [
    ...executiveSummary('The thesis and the evidence for it, in one page.'),
    h1('Introduction'),
    g('State the thesis.'),
    h1('Context'),
    h1('Analysis'),
    h1('Implications and recommendations'),
    h1('Conclusion'),
    h1('References'),
  ],

  memo: () => [
    h1('Purpose'), g('One or two sentences: what this memo asks for or tells the reader.'),
    h1('Background'), g('Only what the reader needs to follow the discussion.'),
    h1('Discussion'), g('The options or the analysis, with the figures that matter.'),
    h1('Recommendation'), g('The decision or action asked for, stated plainly.'),
    h1('Next steps'), actions(['{field:Action, owner and date}']),
  ],

  'blank-report': () => [h1('Executive summary'), h1('Introduction'), h1('{field:Section title}'), h1('Conclusions')],

  'blank-memo': () => [h1('{field:Section title}')],
}
