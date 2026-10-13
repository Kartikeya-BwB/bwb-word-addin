// Consistency: acronyms spelled out on first use (with an abbreviations table for long reports) and
// compounds written more than one way across the document.

import type { AcronymEntry } from '../rules/acronyms.ts'
import { analyseAcronyms } from '../rules/acronyms.ts'
import type { HyphenVariant } from '../rules/hyphenation.ts'
import { findHyphenationVariants } from '../rules/hyphenation.ts'
import type { DocSnapshot, Finding } from './types.ts'
import { excerpt, headingLevel, isTocEntry, wordCount } from './util.ts'

export interface ConsistencyResult {
  findings: Finding[]
  acronyms: AcronymEntry[]
  hyphenation: HyphenVariant[]
}

export function consistencyFindings(snapshot: DocSnapshot): ConsistencyResult {
  const paragraphs = snapshot.paragraphs
  const input = paragraphs.map((p) => ({
    text: isTocEntry(p) ? '' : p.text,
    isHeading: headingLevel(p) !== null,
    inTable: p.inTable,
  }))
  const acronyms = analyseAcronyms(input)
  const findings: Finding[] = []

  for (const a of acronyms) {
    if (a.status !== 'undefined' && a.status !== 'defined-late') continue
    const p = paragraphs[a.firstUse.paragraph]
    findings.push({
      id: `consistency:acronym:${a.acronym}`, rule: `consistency/acronym-${a.status}`, group: 'consistency', severity: 'minor',
      title: a.status === 'undefined' ? `${a.acronym} is never spelled out` : `${a.acronym} is used before it is spelled out`,
      detail: a.status === 'undefined'
        ? 'Spell out acronyms on first use, for example “United Nations Development Programme (UNDP)”.'
        : `It is first spelled out later${a.definition ? ` as “${a.definition}”` : ''}. Move the definition to the first use.`,
      ...(p ? { excerpt: excerpt(p.text, a.firstUse.start, a.firstUse.end), location: { kind: 'paragraph' as const, paragraph: p.index, expectedText: p.text, start: a.firstUse.start, end: a.firstUse.end } } : {}),
    })
  }

  const words = paragraphs.reduce((n, p) => n + wordCount(p.text), 0)
  const counted = acronyms.filter((a) => a.status !== 'common')
  const hasTable = paragraphs.some((p) => headingLevel(p) !== null && /abbreviation|acronym/i.test(p.text))
  if (!hasTable && counted.length >= 8 && words >= 5000) {
    findings.push({
      id: 'consistency:aa-table', rule: 'consistency/abbreviations-table', group: 'consistency', severity: 'minor',
      title: 'No abbreviations table',
      detail: `This document uses ${counted.length} acronyms over ${words.toLocaleString('en-GB')} words. Long reports have an abbreviations and acronyms table; insert one from the Insert tab.`,
    })
  }

  const hyphenation = findHyphenationVariants(paragraphs.filter((p) => !isTocEntry(p)).map((p) => p.text))
  for (const v of hyphenation) {
    const forms = v.forms.map((f) => `“${f.form}” (${f.count})`).join(', ')
    const minority = v.forms.find((f) => f.form !== v.suggestion)
    const p = minority ? paragraphs.find((x) => x.text.toLowerCase().includes(minority.form.toLowerCase())) : undefined
    const start = p && minority ? p.text.toLowerCase().indexOf(minority.form.toLowerCase()) : -1
    findings.push({
      id: `consistency:hyphen:${v.suggestion}`, rule: 'consistency/hyphenation', group: 'consistency', severity: 'copyedit', advisory: true,
      title: `Written more than one way: ${forms}`, detail: `Pick one form. The most common is “${v.suggestion}”.`,
      ...(p && minority && start >= 0
        ? { excerpt: excerpt(p.text, start, start + minority.form.length), location: { kind: 'paragraph' as const, paragraph: p.index, expectedText: p.text, start, end: start + minority.form.length } }
        : {}),
    })
  }
  return { findings, acronyms, hyphenation }
}
