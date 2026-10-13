// Text tools on the selection: changing case word by word (so each word keeps its formatting), non-breaking spaces
// that keep figures together, and inserting symbols.

import { analyseAcronyms } from '../../rules/acronyms.ts'
import { replaceSpans } from '../fixes.ts'
import { UserFacingError, runWord } from '../run.ts'
import { count } from './common.ts'
import type { CaseMode } from './textCase.ts'
import { changeCase, isShouting } from './textCase.ts'
import { EN_DASH, nonBreakingSpans } from './textRules.ts'
import { supportsDesktop } from './support.ts'

const CASE_NAMES: Record<CaseMode, string> = { sentence: 'sentence case', title: 'title case', upper: 'upper case', lower: 'lower case' }

/** Acronyms the document uses in ordinary text, to keep when changing the case of text written in capitals. */
async function documentAcronyms(context: Word.RequestContext): Promise<Set<string>> {
  const paragraphs = context.document.body.paragraphs
  paragraphs.load('items/text,items/tableNestingLevel')
  await context.sync()
  const found = analyseAcronyms(paragraphs.items.map((p) => ({ text: p.text, inTable: p.tableNestingLevel > 0 })))
  return new Set(found.map((a) => a.acronym.toUpperCase()))
}

/**
 * Changes the case of the selected text. Upper case uses Word's own change where it can (WordApiDesktop 1.4); the
 * other cases go word by word so acronyms keep their capitals and each word keeps its formatting.
 */
export async function changeSelectionCase(mode: CaseMode): Promise<string> {
  return runWord(async (context) => {
    const selection = context.document.getSelection()
    selection.load('text')
    const paragraphs = selection.paragraphs
    paragraphs.load('items/text')
    await context.sync()
    if (!selection.text.trim()) throw new UserFacingError('Select the text to change first.')
    if (mode === 'upper' && supportsDesktop('1.4')) {
      selection.case = 'Upper'
      await context.sync()
      return `Selection changed to ${CASE_NAMES[mode]}.`
    }
    const parts = paragraphs.items.map((p) => p.getRange('Content').intersectWithOrNullObject(selection))
    parts.forEach((r) => r.load('isNullObject'))
    await context.sync()
    const words = parts.filter((r) => !r.isNullObject).map((r) => r.getTextRanges([' '], true))
    words.forEach((w) => w.load('items/text'))
    await context.sync()
    const texts = words.map((w) => w.items.map((r) => r.text).join(' '))
    const acronyms = mode !== 'upper' && texts.some(isShouting) ? await documentAcronyms(context) : new Set<string>()
    let changed = 0
    words.forEach((w, i) => {
      const after = changeCase(texts[i] ?? '', mode, { acronyms }).split(' ')
      w.items.forEach((range, k) => {
        const next = after[k]
        if (next !== undefined && next !== range.text) {
          range.insertText(next, 'Replace')
          changed += 1
        }
      })
    })
    await context.sync()
    return changed
      ? `${count(changed, 'word')} changed to ${CASE_NAMES[mode]}.${mode === 'sentence' || mode === 'lower' ? ' Check the names of people, places and organisations.' : ''}`
      : `Nothing to change: the selection is already in ${CASE_NAMES[mode]}.`
  })
}

/**
 * Puts non-breaking spaces into USD 5m, 12 per cent, 5 October 2026 and 20 MW, in the selection or, with nothing
 * selected, in the whole document.
 */
export async function addNonBreakingSpaces(): Promise<string> {
  return runWord(async (context) => {
    const selection = context.document.getSelection()
    selection.load('text')
    await context.sync()
    const whole = !selection.text.trim()
    const paragraphs = whole ? context.document.body.paragraphs : selection.paragraphs
    paragraphs.load('items/text,items/styleBuiltIn')
    await context.sync()
    let places = 0
    let skipped = 0
    for (const p of paragraphs.items) {
      if (/^Toc\d$/.test(p.styleBuiltIn)) continue
      const spans = nonBreakingSpans(p.text)
      if (!spans.length) continue
      const r = await replaceSpans(context, p, p.text, spans)
      places += r.applied
      skipped += r.skipped
    }
    const where = whole ? 'the document' : 'the selection'
    if (!places && !skipped) return `No figures to keep together in ${where}.`
    const left = skipped ? ` ${count(skipped, 'place')} could not be changed safely; check ${skipped === 1 ? 'it' : 'them'} by hand.` : ''
    return `Non-breaking spaces added in ${count(places, 'place')} in ${where}, so figures, dates and units stay on one line.${left}`
  })
}

/** Inserts a symbol at the cursor. The en dash is for ranges in tables only. */
export async function insertSymbol(symbol: string): Promise<void> {
  await runWord(async (context) => {
    const selection = context.document.getSelection()
    if (symbol === EN_DASH) {
      const table = selection.parentTableOrNullObject
      table.load('isNullObject')
      await context.sync()
      if (table.isNullObject) {
        throw new UserFacingError('The en dash is for ranges in tables. In running text, write the range with \u201cto\u201d: 2020 to 2025.')
      }
    }
    selection.insertText(symbol, 'Replace').select('End')
    await context.sync()
  })
}
