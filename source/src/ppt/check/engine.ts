// Runs every deck rule over a snapshot and summarises the result, as the Word check's engine does for documents.

import type { Severity } from '../../check/types.ts'
import { SEVERITIES } from '../../check/types.ts'
import { brandFindings } from './rulesBrand.ts'
import { layoutFindings } from './rulesLayout.ts'
import { sendFindings } from './rulesSend.ts'
import { slideFindings } from './rulesSlides.ts'
import { wordingFindings } from './rulesWording.ts'
import { authoredText, isPicture } from './shapes.ts'
import { wordCount } from './text.ts'
import type { DeckFinding, DeckFix, DeckReport, DeckSnapshot } from './types.ts'

const ORDER: Record<Severity, number> = { blocker: 0, major: 1, minor: 2, copyedit: 3 }

/** Checks this PowerPoint could not run, in plain words for the pane. */
export function deckLimits(snapshot: DeckSnapshot): string[] {
  const out: string[] = []
  if (!snapshot.can.placeholders) {
    out.push('This PowerPoint does not report placeholder types or the shapes inside groups (PowerPoint API 1.8), so titles are found by their names and grouped shapes are not checked.')
  }
  if (!snapshot.can.altText) out.push('Alt text on pictures is not checked: that needs PowerPoint API 1.10.')
  if (!snapshot.can.pageSize) out.push('Shapes off the slide are judged against a 16:9 slide: reading the slide size needs PowerPoint API 1.10.')
  if (!snapshot.can.selection) out.push('Go to needs PowerPoint API 1.5: use the slide numbers instead.')
  if (snapshot.partial) out.push('Some text mixes so much formatting that part of it was not checked for fonts, colours and sizes.')
  return out
}

export function runDeckCheck(snapshot: DeckSnapshot): DeckReport {
  const findings: DeckFinding[] = [
    ...sendFindings(snapshot),
    ...brandFindings(snapshot),
    ...layoutFindings(snapshot),
    ...wordingFindings(snapshot),
    ...slideFindings(snapshot),
  ].sort(
    (a, b) =>
      ORDER[a.severity] - ORDER[b.severity] ||
      Number(Boolean(a.advisory)) - Number(Boolean(b.advisory)) ||
      (a.slide ?? 0) - (b.slide ?? 0),
  )

  const counts = Object.fromEntries(SEVERITIES.map((s) => [s, 0])) as Record<Severity, number>
  for (const f of findings) if (!f.advisory) counts[f.severity] += 1

  const shapes = snapshot.slides.flatMap((s) => s.shapes)
  return {
    findings,
    counts,
    ready: counts.blocker === 0 && counts.major === 0,
    stats: {
      slides: snapshot.slides.length,
      shapes: shapes.length,
      words: snapshot.slides.reduce((n, s) => n + authoredText(s).reduce((m, shape) => m + wordCount(shape.text ?? ''), 0), 0),
      pictures: shapes.filter(isPicture).length,
    },
    limits: deckLimits(snapshot),
  }
}

/** Fixes that change only what the finding names, so they can run in bulk without review. */
const BULK_SAFE = new Set<DeckFix['kind']>(['replace', 'font', 'align-title', 'delete-empty', 'update-trackers'])

export function isBulkSafe(fix: DeckFix): boolean {
  return fix.kind === 'batch' ? fix.fixes.length > 0 && fix.fixes.every(isBulkSafe) : BULK_SAFE.has(fix.kind)
}

export function deckBulkFixes(findings: DeckFinding[]): DeckFinding[] {
  return findings.filter((f) => !f.advisory && f.fix && isBulkSafe(f.fix))
}

/** A short plain-text summary an author can paste when sending the deck to QA. */
export function deckQaNote(report: DeckReport, title: string, date: Date = new Date()): string {
  const day = date.toLocaleDateString('en-GB', { day: 'numeric', month: 'long', year: 'numeric' })
  const c = report.counts
  const open = report.findings.filter((f) => !f.advisory && (f.severity === 'blocker' || f.severity === 'major'))
  const lines = [
    `BwB deck check: ${title || 'untitled deck'}, ${day}`,
    `${report.stats.slides} slides, ${report.stats.words.toLocaleString('en-GB')} words, ${report.stats.pictures} pictures`,
    `Open items: ${c.blocker} blocker, ${c.major} major, ${c.minor} minor, ${c.copyedit} copyedit`,
    report.ready ? 'No blocker or major items: ready for peer review and QA.' : 'Not ready yet. Major items:',
    ...open.slice(0, 10).map((f) => `- ${f.slide ? `Slide ${f.slide}: ` : ''}${f.title}`),
  ]
  return lines.join('\n')
}
