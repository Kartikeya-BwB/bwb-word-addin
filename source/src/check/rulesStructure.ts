// Structure: heading levels and numbering, spacing paragraphs, missing contents, disclaimer and EU wording.

import { checkHeadingCase } from '../rules/headingCase.ts'
import { findDocType } from '../brand/docTypes.ts'
import type { DocSnapshot, Finding, Location, ParagraphInfo } from './types.ts'
import { excerpt, headingLevel, isTocEntry, wordCount } from './util.ts'

// "1.", "1.2", "2.3.1" or "A." at the start of a heading; plain numbers such as "2030 targets" are not matched.
const TYPED_NUMBER = /^\s*(\d{1,2}\.(?:\d{1,2}\.?)*|[A-H]\.|[IVX]{1,4}\.)\s+/
const FRONT_BACK_MATTER = /^(executive summary|summary|key messages|contents|table of contents|glossary|abbreviations|acronyms|references|bibliography|annex|annexes|appendix|appendices|disclaimer)\b/i

function at(p: ParagraphInfo, start?: number, end?: number): Location {
  return { kind: 'paragraph', paragraph: p.index, expectedText: p.text, ...(start !== undefined ? { start } : {}), ...(end !== undefined ? { end } : {}) }
}

function headingFindings(snapshot: DocSnapshot): Finding[] {
  const out: Finding[] = []
  let previous = 0
  for (const p of snapshot.paragraphs) {
    const level = headingLevel(p)
    if (level === null || !p.text.trim()) continue
    if (level > previous + 1 && level <= 4) {
      out.push({
        id: `structure:skip:${p.index}`, rule: 'structure/skipped-level', group: 'structure', severity: 'minor',
        title: `Heading ${level} follows ${previous ? `heading ${previous}` : 'no heading'}`,
        detail: 'Do not skip heading levels: the contents and navigation depend on them.',
        excerpt: excerpt(p.text), location: at(p),
      })
    }
    previous = level

    const typed = TYPED_NUMBER.exec(p.text)
    if (typed && level <= 3) {
      const finding: Finding = {
        id: `structure:typed:${p.index}`, rule: 'structure/typed-number', group: 'structure', severity: 'minor',
        title: 'Heading number typed by hand',
        detail: 'BwB headings number themselves (1., 1.1, 1.1.1). A typed number doubles up or drifts out of order.',
        excerpt: excerpt(p.text), location: at(p, 0, typed[0].length),
      }
      // Only a heading Word already numbers can lose its typed number without losing its number.
      if (p.isListItem) {
        finding.fix = { kind: 'replace', paragraph: p.index, expectedText: p.text, start: 0, end: typed[0].length, match: typed[0], replacement: '' }
        finding.fixLabel = 'Remove the typed number'
      }
      out.push(finding)
    }

    const trimmed = p.text.trimEnd()
    if (trimmed.endsWith('.') && !trimmed.endsWith('etc.') && !trimmed.endsWith('...')) {
      out.push({
        id: `structure:stop:${p.index}`, rule: 'structure/heading-full-stop', group: 'structure', severity: 'copyedit',
        title: 'Full stop at the end of a heading', excerpt: excerpt(p.text), location: at(p, trimmed.length - 1, trimmed.length),
        fix: { kind: 'replace', paragraph: p.index, expectedText: p.text, start: trimmed.length - 1, end: trimmed.length, match: '.', replacement: '' },
        fixLabel: 'Remove the full stop',
      })
    }

    const caseCheck = level <= 3 ? checkHeadingCase(p.text) : { titleCase: false }
    if (caseCheck.titleCase && caseCheck.suggestion && caseCheck.suggestion !== p.text) {
      out.push({
        id: `structure:case:${p.index}`, rule: 'structure/heading-case', group: 'structure', severity: 'copyedit', advisory: true,
        title: 'Heading in title case', detail: 'BwB headings use sentence case. Keep capitals for proper nouns.',
        excerpt: excerpt(p.text), location: at(p),
        fix: { kind: 'replace', paragraph: p.index, expectedText: p.text, start: 0, end: p.text.length, match: p.text, replacement: caseCheck.suggestion },
        fixLabel: `Change to “${caseCheck.suggestion}”`,
      })
    }
  }
  return out
}

function spacingFindings(snapshot: DocSnapshot): Finding[] {
  const out: Finding[] = []
  const ps = snapshot.paragraphs
  for (let i = 1; i < ps.length; i += 1) {
    const p = ps[i]
    const prev = ps[i - 1]
    if (!p || !prev || p.inTable || prev.inTable) continue
    if (p.text === '' && prev.text === '' && !p.isListItem && headingLevel(p) === null) {
      out.push({
        id: `structure:empty:${p.index}`, rule: 'structure/empty-paragraph', group: 'structure', severity: 'copyedit',
        title: 'Empty paragraph used for spacing', detail: 'The BwB styles add the space; extra empty paragraphs make pages drift.',
        location: at(p), fix: { kind: 'delete-paragraph', paragraph: p.index, expectedText: '' }, fixLabel: 'Delete it',
      })
    }
  }
  for (let i = 0; i < ps.length - 1; i += 1) {
    const p = ps[i]
    const next = ps[i + 1]
    if (!p || !next) continue
    const looksLikeHeading =
      headingLevel(p) === null && p.styleBuiltIn === 'Normal' && !p.inTable && !p.isListItem && p.font.bold === true &&
      p.text.trim().length >= 3 && wordCount(p.text) <= 10 && !/[.:;,?!]$/.test(p.text.trim()) && next.text.trim() !== ''
    if (looksLikeHeading) {
      out.push({
        id: `structure:fake:${p.index}`, rule: 'structure/fake-heading', group: 'structure', severity: 'minor',
        title: 'Bold text used as a heading',
        detail: 'Apply a heading style so it is numbered and appears in the contents and the navigation pane.',
        excerpt: excerpt(p.text), location: at(p),
      })
    }
  }
  for (const p of ps) {
    if (headingLevel(p) === null && !p.inTable && wordCount(p.text) > 200) {
      out.push({
        id: `structure:long:${p.index}`, rule: 'structure/long-paragraph', group: 'structure', severity: 'copyedit', advisory: true,
        title: `Long paragraph (${wordCount(p.text)} words)`, detail: 'Consider splitting it: one idea per paragraph.',
        excerpt: excerpt(p.text, 0, 60), location: at(p),
      })
    }
  }
  return out
}

function documentFindings(snapshot: DocSnapshot): Finding[] {
  const out: Finding[] = []
  const ps = snapshot.paragraphs
  const docType = findDocType(snapshot.info.docType)

  const mainHeadings = ps.filter((p) => {
    const l = headingLevel(p)
    return l !== null && l <= 2 && p.text.trim() !== ''
  })
  const hasToc = (snapshot.fieldCodes ?? []).some((c) => /^\s*TOC\b/i.test(c)) || ps.some(isTocEntry)
  const isReport = docType ? docType.layout === 'report' : ps.some((p) => p.styleBuiltIn === 'Title')
  if (isReport && mainHeadings.length >= 5 && !hasToc) {
    out.push({
      id: 'structure:toc', rule: 'structure/no-contents', group: 'structure', severity: 'minor',
      title: 'No table of contents',
      detail: `Reports with five or more main headings (this one has ${mainHeadings.length}) have contents after the cover. Insert one from the Insert tab.`,
    })
  }

  const hasDisclaimer = ps.some((p) => p.style === 'BwB Disclaimer Heading' || p.text.includes('FRN: 953637') || /^disclaimer$/i.test(p.text.trim()))
  const needsDisclaimer = docType ? docType.disclaimer : isReport
  if (needsDisclaimer && !hasDisclaimer) {
    out.push({
      id: 'structure:disclaimer', rule: 'structure/no-disclaimer', group: 'send', severity: docType ? 'major' : 'minor',
      title: 'No BwB disclaimer',
      detail: 'Client-facing documents end with the standard disclaimer. Check the variant in Disclaimers v1.0 for investment-related material.',
      fix: { kind: 'insert-disclaimer' }, fixLabel: 'Add it at the end',
    })
  }

  if (snapshot.info.funder === 'eu-grant' && !ps.some((p) => /funded by the european union/i.test(p.text))) {
    out.push({
      id: 'structure:eu', rule: 'structure/eu-acknowledgement', group: 'send', severity: 'major',
      title: 'No EU funding acknowledgement',
      detail: 'EU grant deliverables carry the EU emblem, the funding acknowledgement and the EU disclaimer. Insert it from the Insert tab and check the wording in the grant agreement.',
    })
  }

  // Word reports a heading numbered through its style's multilevel list as a list item.
  const numberedFrontMatter = ps.filter((p) => {
    const l = headingLevel(p)
    return l === 1 && p.styleBuiltIn === 'Heading1' && p.isListItem && FRONT_BACK_MATTER.test(p.text.trim())
  })
  for (const p of numberedFrontMatter) {
    out.push({
      id: `structure:front:${p.index}`, rule: 'structure/numbered-front-matter', group: 'structure', severity: 'copyedit',
      title: 'Front or back matter is numbered',
      detail: 'Executive summary, glossary, references and annexes are not numbered.',
      excerpt: excerpt(p.text), location: at(p),
      fix: { kind: 'paragraph-style', paragraph: p.index, expectedText: p.text, style: 'BwB Heading 1 Unnumbered' },
      fixLabel: 'Make it an unnumbered heading',
    })
  }
  return out
}

export function structureFindings(snapshot: DocSnapshot): Finding[] {
  return [...headingFindings(snapshot), ...spacingFindings(snapshot), ...documentFindings(snapshot)]
}
