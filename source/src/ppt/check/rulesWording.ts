// House-style wording on slides, from the same rules as the Word check (src/rules/writing.ts), run on each
// paragraph of each shape. Quotations, BwB furniture and footers are left alone. Placeholder text left in is a
// "before you send" item.

import type { WritingCode, WritingIssue } from '../../rules/writing.ts'
import { quoteIssues, scan } from '../../rules/writing.ts'
import { BODY_PLACEHOLDERS, hasText, isBwbFurniture, isFooterShape, isQuoteShape, isTitleShape } from './shapes.ts'
import { splitParagraphs } from './text.ts'
import type { DeckFinding, DeckSnapshot, ShapeInfo, SlideInfo } from './types.ts'
import { excerpt, onShape, shapeKey } from './util.ts'

export const WORDING_TITLES: Record<WritingCode, string> = {
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

function toFinding(issue: WritingIssue, slide: SlideInfo, shape: ShapeInfo, offset: number): DeckFinding {
  const text = shape.text ?? ''
  const start = offset + issue.start
  const end = offset + issue.end
  const placeholder = issue.code === 'placeholder'
  const finding: DeckFinding = {
    id: `wording:${issue.code}:${shapeKey(slide, shape)}:${start}`,
    rule: `wording/${issue.code}`,
    group: placeholder ? 'send' : 'wording',
    severity: placeholder ? 'major' : 'copyedit',
    title: WORDING_TITLES[issue.code],
    detail: typographic(issue.message),
    excerpt: excerpt(text, start, end),
    slide: slide.index + 1,
    location: onShape(slide, shape, start, end),
  }
  if (issue.severity === 'info' && !placeholder) finding.advisory = true
  if (issue.fixable && issue.replacement !== undefined) {
    finding.fix = {
      kind: 'replace', slideId: slide.id, path: shape.path, expectedText: text, start, end, match: issue.match, replacement: issue.replacement,
    }
    finding.fixLabel = issue.replacement.trim() ? `Change to “${issue.replacement.trim()}”` : 'Remove'
  }
  return finding
}

export function wordingFindings(snapshot: DeckSnapshot): DeckFinding[] {
  const out: DeckFinding[] = []
  for (const slide of snapshot.slides) {
    for (const shape of slide.shapes) {
      if (!hasText(shape) || isBwbFurniture(shape) || isFooterShape(shape) || isQuoteShape(shape)) continue
      const title = isTitleShape(shape)
      // Labels, callouts and table-like text boxes are compact, as table cells are in Word; body placeholders are prose.
      const compact = !title && !(shape.placeholder !== null && BODY_PLACEHOLDERS.has(shape.placeholder))
      for (const p of splitParagraphs(shape.text ?? '')) {
        if (!p.text.trim()) continue
        for (const issue of [...scan(p.text, { title, compact }), ...quoteIssues(p.text)]) out.push(toFinding(issue, slide, shape, p.start))
      }
    }
  }
  return out
}
