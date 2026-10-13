// Document parts: contents, the disclaimer, the EU acknowledgement and the abbreviations table.

import { euAcknowledgementBlocks } from '../brand/eu.ts'
import { consistencyFindings } from '../check/rulesConsistency.ts'
import { abbreviationsBlocks } from './abbreviations.ts'
import { isTocField, updateFields } from './fields.ts'
import { insertBlocks } from './insert.ts'
import { UserFacingError, runWord, supports } from './run.ts'
import { takeSnapshot } from './snapshot.ts'

const withWarnings = (done: string, warnings: string[]): string => (warnings.length ? `${done} ${warnings.join(' ')}` : done)

/** The regulatory line every version of the BwB disclaimer ends with. */
const DISCLAIMER_MARK = '953637'

async function bodyContains(text: string): Promise<boolean> {
  return runWord(async (context) => {
    const hits = context.document.body.search(text, { matchCase: true })
    hits.load('items')
    await context.sync()
    return hits.items.length > 0
  })
}

export async function insertDisclaimer(): Promise<string> {
  if (await bodyContains(DISCLAIMER_MARK)) {
    throw new UserFacingError('This document already has the BwB disclaimer: it names FRN 953637.')
  }
  const { warnings } = await insertBlocks([{ type: 'disclaimer' }], 'end')
  return withWarnings('Disclaimer added on its own page at the end.', warnings)
}

export async function insertContents(): Promise<string> {
  const { warnings } = await insertBlocks([{ type: 'toc' }])
  let filled = false
  if (supports('1.5')) {
    try {
      filled = (await updateFields(isTocField)) > 0
    } catch {
      filled = false
    }
  }
  return withWarnings(filled ? 'Contents inserted and filled.' : 'Contents inserted. Update fields (F9) to fill it.', warnings)
}

export async function insertEuAcknowledgement(): Promise<string> {
  const { warnings } = await insertBlocks(euAcknowledgementBlocks())
  return withWarnings('EU acknowledgement inserted. Check the wording against the grant agreement.', warnings)
}

export async function insertAbbreviations(): Promise<string> {
  const { acronyms } = consistencyFindings(await takeSnapshot())
  if (!acronyms.length) throw new UserFacingError('No acronyms were found in this document.')
  const { warnings } = await insertBlocks(abbreviationsBlocks(acronyms))
  const missing = acronyms.filter((a) => !a.definition && a.status !== 'common').length
  const note = missing ? ` ${missing} meaning${missing === 1 ? ' is' : 's are'} not given in the text: fill in the placeholders.` : ''
  return withWarnings(`Abbreviations table inserted with ${acronyms.length} entries.${note}`, warnings)
}
