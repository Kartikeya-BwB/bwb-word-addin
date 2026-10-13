// Author checklists. Wording follows the BwB Deliverables Hub plan (QA Policy v2.0 and Style Guide v2.0, May 2026)
// and the document-type research of 10 October 2026. Items with an `auto` key are also checked by the Pre-QA check.

export type AutoCheck =
  | 'spelling'
  | 'consistency'
  | 'acronyms'
  | 'captions'
  | 'properties'
  | 'placeholders'
  | 'styles'
  | 'toc'
  | 'header-footer'
  | 'tracked-changes'
  | 'alt-text'
  | 'file-name'

export interface ChecklistItem {
  id: string
  text: string
  auto?: AutoCheck
}

export const AUTHOR_SELF_CHECK: ChecklistItem[] = [
  { id: 'objectives', text: 'Objectives, audience and client agreement reviewed' },
  { id: 'template', text: 'Correct template used (client or BwB)', auto: 'styles' },
  { id: 'spelling', text: 'British spelling, unless the client requires otherwise', auto: 'spelling' },
  { id: 'consistency', text: 'Headings, lists and spacing consistent', auto: 'consistency' },
  { id: 'acronyms', text: 'Acronyms spelled out on first use, with an abbreviations table for long reports', auto: 'acronyms' },
  { id: 'facts', text: 'Facts, figures and claims checked against the latest data: one number, one place' },
  { id: 'captions', text: 'Tables and figures labelled and sourced', auto: 'captions' },
  { id: 'references', text: 'References complete before peer review' },
  { id: 'logos', text: 'Third-party logo permissions documented' },
  { id: 'links', text: 'Hyperlinks work' },
  { id: 'title', text: 'Title and authorship correct', auto: 'properties' },
  { id: 'placeholders', text: 'No placeholders ("TBC", "[insert]"), highlights or text from another client', auto: 'placeholders' },
  { id: 'ai-review', text: 'AI review skill run' },
  { id: 'ai-use', text: 'AI use disclosed; no unedited AI text sent for review' },
]

export const WORD_FILE_CHECKS: ChecklistItem[] = [
  { id: 'docx-styles', text: 'Template styles used', auto: 'styles' },
  { id: 'docx-toc', text: 'Table of contents and cross-references updated', auto: 'toc' },
  { id: 'docx-header-footer', text: 'Headers and footers present, page numbers bottom right', auto: 'header-footer' },
  { id: 'docx-tracked', text: 'No tracked changes or comments', auto: 'tracked-changes' },
  { id: 'docx-properties', text: 'Document properties (author, company, title) cleaned', auto: 'properties' },
  { id: 'docx-alt', text: 'Alt text on figures', auto: 'alt-text' },
  { id: 'docx-name', text: 'File name has no internal markers (Copy, OLD, initials, internal only, FINAL on a draft)', auto: 'file-name' },
]

export const EU_GRANT_CHECKS: ChecklistItem[] = [
  { id: 'eu-template', text: 'Consortium template and cover page: title, code, authors, dissemination level and version table' },
  { id: 'eu-emblem', text: 'EU emblem and funding acknowledgement included' },
  { id: 'eu-disclaimer', text: 'EU disclaimer included' },
  { id: 'eu-code', text: 'Deliverable code, title and month identical everywhere they appear' },
  { id: 'eu-ga', text: 'Content matches the grant agreement description; no double reporting' },
  { id: 'eu-sen', text: 'A SEN (sensitive) classification has a stated rationale' },
  { id: 'eu-partners', text: 'Partner contributions attributed' },
  { id: 'eu-entity', text: 'Legal entity is BwB Connect CLG where applicable' },
]

export const COMPLIANCE_CHECKS: ChecklistItem[] = [
  { id: 'fca-review', text: 'Possible financial promotion: Compliance Lead review and co-sign arranged' },
  { id: 'fca-disclaimer', text: 'The correct disclaimer variant is used (Disclaimers v1.0)' },
  { id: 'fca-advice', text: 'No implied investment recommendation; statements framed as BwB analysis, not advice' },
]
