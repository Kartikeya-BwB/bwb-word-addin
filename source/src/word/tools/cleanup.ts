// Tidying pasted text in the selection: rejoining lines broken by a PDF, clearing stray character formatting, and
// removing empty paragraphs. Each edits Word's text in place, so the formatting of the words themselves survives.

import { UserFacingError, runWord } from '../run.ts'
import { count, listOf } from './common.ts'
import { fragmentRisk, isRemovableEmpty, stripRunFormatting } from './ooxml.ts'
import type { LineBreak } from './textRules.ts'
import { planJoins } from './textRules.ts'

interface Lookup {
  breaks: LineBreak[]
  lines: Word.RangeCollection | null
  hyphenLines: Word.RangeCollection | null
  end: Word.RangeCollection | null
}

const occurrences = (text: string, needle: string): number => text.split(needle).length - 1

/**
 * Rejoins lines pasted from a PDF: a paragraph mark or manual line break that ends a line mid-sentence becomes a
 * space, a word hyphenated at the end of a line is joined up again, and double spaces are collapsed.
 */
export async function joinBrokenLines(): Promise<string> {
  return runWord(async (context) => {
    const paragraphs = context.document.getSelection().paragraphs
    paragraphs.load('items/text,items/tableNestingLevel')
    await context.sync()
    const items = paragraphs.items
    if (items.some((p) => p.tableNestingLevel > 0)) throw new UserFacingError('Select text outside tables.')
    if (items.length < 2 && !items.some((p) => p.text.includes('\u000b'))) throw new UserFacingError('Select the lines to join first.')
    const plan = planJoins(items.map((p) => p.text))
    const first = items[0]
    const last = items[items.length - 1]
    if (!first || !last) throw new UserFacingError('Select the lines to join first.')
    const span = first.getRange('Whole').expandTo(last.getRange('Whole'))

    const lookups: Lookup[] = items.map((p, i) => {
      const breaks = plan.filter((b) => b.paragraph === i && b.action !== 'keep')
      const lineBreaks = breaks.filter((b) => b.kind === 'line')
      const end = breaks.find((b) => b.kind === 'paragraph')
      return {
        breaks,
        lines: lineBreaks.length ? p.search('^l', { matchCase: true }) : null,
        hyphenLines: lineBreaks.some((b) => b.action === 'dehyphenate') ? p.search('-^l', { matchCase: true }) : null,
        end: end ? p.getRange('Whole').search(end.action === 'dehyphenate' ? '-^p' : '^p', { matchCase: true }) : null,
      }
    })
    lookups.forEach((l) => [l.lines, l.hyphenLines, l.end].forEach((c) => c?.load('items')))
    await context.sync()

    let joined = 0
    let dehyphenated = 0
    let skipped = 0
    const replace = (range: Word.Range | undefined, b: LineBreak): void => {
      if (!range) {
        skipped += 1
        return
      }
      range.insertText(b.action === 'space' ? ' ' : '', 'Replace')
      if (b.action === 'dehyphenate') dehyphenated += 1
      else joined += 1
    }
    // From the end of the selection back, so each break is still where the search found it.
    for (let i = items.length - 1; i >= 0; i -= 1) {
      const l = lookups[i]
      const text = items[i]?.text ?? ''
      if (!l) continue
      const end = l.breaks.find((b) => b.kind === 'paragraph')
      if (end) replace(l.end?.items.length === 1 ? l.end.items[0] : undefined, end)
      const lineBreaks = l.breaks.filter((b) => b.kind === 'line').sort((a, b) => b.index - a.index)
      const lined = l.lines?.items.length === occurrences(text, '\u000b')
      const hyphened = !l.hyphenLines || l.hyphenLines.items.length === occurrences(text, '-\u000b')
      for (const b of lineBreaks) {
        if (!lined || !hyphened) replace(undefined, b)
        else replace(b.action === 'dehyphenate' ? l.hyphenLines?.items[b.hyphenIndex] : l.lines?.items[b.index], b)
      }
    }
    await context.sync()

    let collapsed = 0
    try {
      for (let round = 0; round < 4; round += 1) {
        const doubles = span.search('  ', { matchCase: true })
        doubles.load('items')
        await context.sync()
        if (!doubles.items.length) break
        for (const d of [...doubles.items].reverse()) d.insertText(' ', 'Replace')
        collapsed += doubles.items.length
        await context.sync()
      }
    } catch {
      // The joins are done; double spaces left over can be cleared with Clean up wording.
    }
    if (!joined && !dehyphenated && !collapsed) {
      return skipped ? `${count(skipped, 'line break')} could not be joined safely: join ${skipped === 1 ? 'it' : 'them'} by hand.` : 'No broken lines in the selection.'
    }
    const done = [
      joined ? `${count(joined, 'broken line')} joined` : '',
      dehyphenated ? `${count(dehyphenated, 'hyphenated word')} rejoined` : '',
      collapsed ? `${count(collapsed, 'double space')} removed` : '',
    ].filter(Boolean)
    const left = skipped ? ` ${count(skipped, 'line break')} could not be joined safely.` : ''
    const sentence = listOf(done)
    return `${sentence.charAt(0).toUpperCase()}${sentence.slice(1)}.${left}`
  })
}

/**
 * Removes stray character formatting (fonts, sizes, colours, highlighting) from the selection, or from the paragraph
 * at the cursor, so the BwB styles show. Bold, italic, underline, links and superscripts stay. It rewrites the runs
 * through WordprocessingML, which works the same on the web and the desktop; paragraphs holding tracked changes,
 * comments, notes, bookmarks or tables are left alone rather than risk losing them.
 */
export async function clearStrayFormatting(): Promise<string> {
  return runWord(async (context) => {
    const selection = context.document.getSelection()
    selection.load('text')
    const paragraphs = selection.paragraphs
    paragraphs.load('items/text')
    await context.sync()
    const whole = !selection.text.trim()
    const withText = paragraphs.items.filter((p) => p.text.trim())
    if (!withText.length) throw new UserFacingError('Select the text to clear first.')
    const parts = withText.map((p) => (whole ? p.getRange('Content') : p.getRange('Content').intersectWithOrNullObject(selection)))
    parts.forEach((r) => r.load('isNullObject'))
    await context.sync()
    const live = parts.filter((r) => !r.isNullObject)
    const xml = live.map((r) => r.getOoxml())
    await context.sync()
    let cleared = 0
    const risks = new Map<string, number>()
    live.forEach((range, i) => {
      const fragment = xml[i]?.value ?? ''
      const risk = fragmentRisk(fragment)
      if (risk) {
        risks.set(risk, (risks.get(risk) ?? 0) + 1)
        return
      }
      const stripped = stripRunFormatting(fragment)
      if (!stripped.removed) return
      range.insertOoxml(stripped.xml, 'Replace')
      cleared += 1
    })
    await context.sync()
    const leftCount = [...risks.values()].reduce((a, n) => a + n, 0)
    const left = leftCount
      ? ` ${count(leftCount, 'paragraph was', 'paragraphs were')} left alone so ${leftCount === 1 ? 'its' : 'their'} ${listOf([...risks.keys()])} stay intact.`
      : ''
    if (!cleared) return `No stray formatting to clear.${left}`
    return `Stray formatting cleared from ${count(cleared, 'paragraph')}. Bold, italic and underlining were kept.${left}`
  })
}

/** Removes the empty paragraphs in the selection, keeping any that hold a picture, a break or anything else. */
export async function removeEmptyParagraphs(): Promise<string> {
  return runWord(async (context) => {
    const paragraphs = context.document.getSelection().paragraphs
    paragraphs.load('items/text,items/tableNestingLevel,items/isLastParagraph')
    await context.sync()
    const candidates = paragraphs.items.filter((p) => !p.text.replace(/[ \u00a0\u200b]/g, '') && p.tableNestingLevel === 0 && !p.isLastParagraph)
    if (!candidates.length) return 'No empty paragraphs in the selection.'
    const checks = candidates.map((p) => {
      const before = p.getPreviousOrNullObject()
      const after = p.getNextOrNullObject()
      before.load('tableNestingLevel')
      after.load('tableNestingLevel')
      return { p, before, after, xml: p.getOoxml() }
    })
    await context.sync()
    // An empty paragraph between two tables keeps them apart: without it Word joins them.
    const between = (c: (typeof checks)[number]): boolean =>
      !c.before.isNullObject && !c.after.isNullObject && c.before.tableNestingLevel > 0 && c.after.tableNestingLevel > 0
    const removable = checks.filter((c) => !between(c) && isRemovableEmpty(c.xml.value))
    for (const c of [...removable].reverse()) c.p.delete()
    await context.sync()
    const kept = candidates.length - removable.length
    const note = kept ? ` ${count(kept, 'paragraph')} that ${kept === 1 ? 'looks' : 'look'} empty but ${kept === 1 ? 'holds' : 'hold'} a picture, a break or a section break, or ${kept === 1 ? 'separates' : 'separate'} two tables, ${kept === 1 ? 'was' : 'were'} kept.` : ''
    return removable.length ? `${count(removable.length, 'empty paragraph')} removed.${note}` : `Nothing removed.${note}`
  })
}
