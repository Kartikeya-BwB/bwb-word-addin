// House-style wording, from the same rules the BwB Document Format skill applies (ported in src/rules).

import type { WritingCode, WritingIssue } from '../rules/writing.ts'
import { quoteIssues, scan } from '../rules/writing.ts'
import type { DocSnapshot, Finding, Severity } from './types.ts'
import { excerpt, headingLevel, isTitle, isVerbatim } from './util.ts'

const TITLES: Record<WritingCode, string> = {
  'em-dash': 'Em dash',
  'en-dash-range': 'En dash in a range',
  'en-dash': 'Spaced en dash',
  'en-dash-compound': 'En dash between words',
  'range-hyphen': 'Year range with a hyphen',
  'per-cent': '“percent”',
  'us-spelling': 'US spelling',
  'us-spelling-check': 'Possible US spelling',
  'date-format': 'Date format',
  'date-numeric': 'Numeric date',
  currency: 'Currency symbol',
  'currency-symbol': 'Currency symbol',
  'currency-space': 'Currency code spacing',
  'brand-name': 'Company name',
  'double-space': 'Double space',
  placeholder: 'Placeholder text left in',
  filler: 'Filler phrase',
  quotes: 'Straight quotation mark',
}

/** The rules' messages quote with straight single quotes, as the Python does; the pane shows typographic ones. */
const typographic = (message: string): string => message.replace(/'([^'\n]*)'/g, '‘$1’')

function toFinding(issue: WritingIssue, paragraph: number, text: string): Finding {
  const placeholder = issue.code === 'placeholder'
  const severity: Severity = placeholder ? 'major' : 'copyedit'
  const finding: Finding = {
    id: `writing:${issue.code}:${paragraph}:${issue.start}`,
    rule: `writing/${issue.code}`,
    group: placeholder ? 'send' : 'writing',
    severity,
    title: TITLES[issue.code],
    detail: typographic(issue.message),
    excerpt: excerpt(text, issue.start, issue.end),
    location: { kind: 'paragraph', paragraph, expectedText: text, start: issue.start, end: issue.end },
  }
  if (issue.severity === 'info' && !placeholder) finding.advisory = true
  if (issue.fixable && issue.replacement !== undefined) {
    finding.fix = {
      kind: 'replace',
      paragraph,
      expectedText: text,
      start: issue.start,
      end: issue.end,
      match: issue.match,
      replacement: issue.replacement,
    }
    finding.fixLabel = issue.replacement.trim() ? `Change to “${issue.replacement.trim()}”` : 'Remove'
  }
  return finding
}

export function writingFindings(snapshot: DocSnapshot): Finding[] {
  const guidance = new Set(snapshot.controls.filter((c) => c.tag === 'BwBGuidance').map((c) => c.text.trim()))
  // Word leaves inline pictures out of a paragraph's text, so the spaces either side of an icon read as a double space.
  const withPictures = new Set((snapshot.formatting ?? []).filter((f) => f.pictures > 0 && f.paragraph !== null).map((f) => f.paragraph))
  const out: Finding[] = []
  for (const p of snapshot.paragraphs) {
    if (!p.text.trim() || isVerbatim(p) || guidance.has(p.text.trim())) continue
    const ctx = { title: headingLevel(p) !== null || isTitle(p), compact: p.inTable }
    for (const issue of scan(p.text, ctx)) {
      if (issue.code === 'double-space' && withPictures.has(p.index)) continue
      out.push(toFinding(issue, p.index, p.text))
    }
    for (const issue of quoteIssues(p.text)) out.push(toFinding(issue, p.index, p.text))
  }
  return out
}
