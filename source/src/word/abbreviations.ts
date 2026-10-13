// The abbreviations and acronyms table long BwB reports carry in their front matter.

import type { AcronymEntry } from '../rules/acronyms.ts'
import type { Block } from './docxApi.ts'

/** Meanings of the acronyms the checker treats as common, which documents rarely spell out. */
const COMMON_MEANINGS: Record<string, string> = {
  UK: 'United Kingdom',
  US: 'United States',
  EU: 'European Union',
  UN: 'United Nations',
  BwB: 'Bankers without Boundaries',
}

export const ABBREVIATIONS_HEADING = 'Abbreviations and acronyms'

/** An unnumbered heading and a two-column table, with a placeholder for any meaning the document never gives. */
export function abbreviationsBlocks(acronyms: AcronymEntry[]): Block[] {
  const rows = [...acronyms]
    .sort((a, b) => a.acronym.localeCompare(b.acronym, 'en-GB', { sensitivity: 'base' }))
    .map((a) => [a.acronym, a.definition ?? COMMON_MEANINGS[a.acronym] ?? `{field:Meaning of ${a.acronym}}`])
  return [
    { type: 'heading', text: ABBREVIATIONS_HEADING, level: 1, numbered: false, in_toc: true },
    { type: 'table', header: ['Abbreviation', 'Meaning'], rows, widths: [1, 4] },
  ]
}
