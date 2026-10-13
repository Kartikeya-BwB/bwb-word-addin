// Runs every Pre-QA rule over a document snapshot and summarises the result.

import { brandFindings } from './rulesBrand.ts'
import { consistencyFindings } from './rulesConsistency.ts'
import { sendFindings } from './rulesSend.ts'
import { structureFindings } from './rulesStructure.ts'
import { tableFindings } from './rulesTables.ts'
import { writingFindings } from './rulesWriting.ts'
import type { CheckReport, DocSnapshot, Finding, Severity } from './types.ts'
import { SEVERITIES } from './types.ts'
import { headingLevel, wordCount } from './util.ts'

const ORDER: Record<Severity, number> = { blocker: 0, major: 1, minor: 2, copyedit: 3 }

export function runChecks(snapshot: DocSnapshot): CheckReport {
  const consistency = consistencyFindings(snapshot)
  const findings: Finding[] = [
    ...sendFindings(snapshot),
    ...structureFindings(snapshot),
    ...tableFindings(snapshot),
    ...brandFindings(snapshot),
    ...consistency.findings,
    ...writingFindings(snapshot),
  ].sort((a, b) => ORDER[a.severity] - ORDER[b.severity] || Number(Boolean(a.advisory)) - Number(Boolean(b.advisory)))

  const counts = Object.fromEntries(SEVERITIES.map((s) => [s, 0])) as Record<Severity, number>
  for (const f of findings) if (!f.advisory) counts[f.severity] += 1

  return {
    findings,
    counts,
    ready: counts.blocker === 0 && counts.major === 0,
    acronyms: consistency.acronyms,
    hyphenation: consistency.hyphenation,
    stats: {
      paragraphs: snapshot.paragraphs.length,
      words: snapshot.paragraphs.reduce((n, p) => n + wordCount(p.text), 0),
      headings: snapshot.paragraphs.filter((p) => headingLevel(p) !== null && p.text.trim() !== '').length,
      tables: snapshot.tables.length,
      pictures: snapshot.pictures.length,
    },
  }
}

/** A short plain-text summary an author can paste when sending the document to QA. */
export function qaNote(report: CheckReport, title: string, date: Date = new Date()): string {
  const day = date.toLocaleDateString('en-GB', { day: 'numeric', month: 'long', year: 'numeric' })
  const c = report.counts
  const open = report.findings.filter((f) => !f.advisory && (f.severity === 'blocker' || f.severity === 'major'))
  const lines = [
    `BwB Pre-QA check: ${title || 'untitled document'}, ${day}`,
    `${report.stats.words.toLocaleString('en-GB')} words, ${report.stats.headings} headings, ${report.stats.tables} tables`,
    `Open items: ${c.blocker} blocker, ${c.major} major, ${c.minor} minor, ${c.copyedit} copyedit`,
    report.ready ? 'No blocker or major items: ready for peer review and QA.' : 'Not ready yet. Major items:',
    ...open.slice(0, 10).map((f) => `- ${f.title}`),
  ]
  return lines.join('\n')
}
