// Applying BwB styles, and bringing the BwB styles into documents that were not built on the template.

import type { BuiltIn, ListStyle, ParagraphStyle } from '../brand/styles.ts'
import { TEMPLATE_MARKER_STYLES } from '../brand/styles.ts'
import { bodyToFlatOpc, styleCarrierBody } from './flatOpc.ts'
import { runWord, supports } from './run.ts'

export type StyleTarget = Pick<ParagraphStyle, 'name'> & { builtIn?: BuiltIn }

/** Null when this Word cannot tell (WordApi 1.5 is needed to list styles). */
export async function hasBwbStyles(): Promise<boolean | null> {
  if (!supports('1.5')) return null
  return runWord(async (context) => {
    const styles = context.document.getStyles()
    const found = TEMPLATE_MARKER_STYLES.map((name) => styles.getByNameOrNullObject(name))
    found.forEach((s) => s.load('isNullObject'))
    await context.sync()
    return found.every((s) => !s.isNullObject)
  })
}

/**
 * Copies every custom BwB style (and the list definitions they use) into the open document by inserting
 * content that uses them at the end of the document and deleting it again. Styles the document already
 * has keep their current definition: Word never overwrites a destination style on insert.
 */
export async function addBwbStyles(): Promise<void> {
  await runWord(async (context) => {
    const body = context.document.body
    const before = body.paragraphs
    before.load('items/text')
    await context.sync()
    const count = before.items.length

    const inserted = body.insertOoxml(bodyToFlatOpc(styleCarrierBody()), 'End')
    await context.sync()
    inserted.delete()
    await context.sync()

    // Inserting at the end can leave one extra empty paragraph behind; remove it so the text is unchanged.
    const after = body.paragraphs
    after.load('items/text')
    await context.sync()
    const last = after.items[after.items.length - 1]
    if (after.items.length > count && last && last.text === '') {
      last.delete()
      await context.sync()
    }
  })
}

interface BuiltInSpec {
  name: string
  font: Word.Interfaces.FontUpdateData
  paragraph: Word.Interfaces.ParagraphFormatUpdateData
}

const NAVY = '#022338'

const BUILT_IN_SPECS: BuiltInSpec[] = [
  { name: 'Normal', font: { name: 'Aptos', size: 10, color: '#000000' }, paragraph: { alignment: 'Justified', spaceAfter: 8, lineSpacing: 16 } },
  { name: 'Heading 1', font: { name: 'Candara', size: 18, color: NAVY, bold: true, italic: false }, paragraph: { alignment: 'Left', spaceBefore: 12, spaceAfter: 6, keepWithNext: true, keepTogether: true } },
  { name: 'Heading 2', font: { name: 'Candara', size: 14, color: NAVY, bold: true, italic: false }, paragraph: { alignment: 'Left', spaceBefore: 10, spaceAfter: 4, keepWithNext: true, keepTogether: true } },
  { name: 'Heading 3', font: { name: 'Candara', size: 12, color: NAVY, bold: true, italic: false }, paragraph: { alignment: 'Left', spaceBefore: 8, spaceAfter: 3, keepWithNext: true, keepTogether: true } },
  { name: 'Heading 4', font: { name: 'Candara', size: 11, color: NAVY, bold: true, italic: true }, paragraph: { alignment: 'Left', spaceBefore: 6, spaceAfter: 2, keepWithNext: true, keepTogether: true } },
  { name: 'Caption', font: { name: 'Aptos', size: 9, color: '#4A4A4C', italic: true, bold: false }, paragraph: { alignment: 'Left', spaceBefore: 3, spaceAfter: 10 } },
]

/**
 * Restyles Word's own styles (Normal, Heading 1 to 4, Caption) to the BwB specification so documents
 * that did not start from the template look right. Needs WordApi 1.5. Returns the styles it changed.
 */
export async function restyleBuiltInStyles(): Promise<string[]> {
  return runWord(async (context) => {
    const styles = context.document.getStyles()
    const found = BUILT_IN_SPECS.map((spec) => ({ spec, style: styles.getByNameOrNullObject(spec.name) }))
    found.forEach(({ style }) => style.load('isNullObject'))
    await context.sync()
    const changed: string[] = []
    for (const { spec, style } of found) {
      if (style.isNullObject) continue
      style.font.set(spec.font)
      style.paragraphFormat.set(spec.paragraph)
      changed.push(spec.name)
    }
    await context.sync()
    return changed
  })
}

async function styleExists(context: Word.RequestContext, name: string): Promise<boolean | null> {
  if (!supports('1.5')) return null
  const style = context.document.getStyles().getByNameOrNullObject(name)
  style.load('isNullObject')
  await context.sync()
  return !style.isNullObject
}

function setStyle(p: Word.Paragraph, style: StyleTarget): void {
  if (style.builtIn) p.styleBuiltIn = style.builtIn
  else p.style = style.name
}

export interface ApplyResult {
  paragraphs: number
  importedStyles: boolean
}

/**
 * Runs a change that uses a custom BwB style, bringing the BwB styles into the document first when it lacks
 * that style, so one click works in any document.
 */
export async function withBwbStyle<T>(name: string, apply: () => Promise<T>): Promise<{ result: T; importedStyles: boolean }> {
  let importedStyles = false
  if ((await runWord((context) => styleExists(context, name))) === false) {
    await addBwbStyles()
    importedStyles = true
  }
  try {
    return { result: await apply(), importedStyles }
  } catch (error) {
    // Older Word cannot list styles, so a missing style only shows up as a failed apply.
    if (importedStyles) throw error
    await addBwbStyles()
    return { result: await apply(), importedStyles: true }
  }
}

/** Applies a paragraph style to every paragraph the selection touches. */
export async function applyParagraphStyle(style: StyleTarget): Promise<ApplyResult> {
  const apply = (): Promise<number> =>
    runWord(async (context) => {
      const paragraphs = context.document.getSelection().paragraphs
      paragraphs.load('items/text')
      await context.sync()
      paragraphs.items.forEach((p) => setStyle(p, style))
      await context.sync()
      return paragraphs.items.length
    })
  if (style.builtIn) return { paragraphs: await apply(), importedStyles: false }
  const { result, importedStyles } = await withBwbStyle(style.name, apply)
  return { paragraphs: result, importedStyles }
}

export function applyListStyle(list: ListStyle): Promise<ApplyResult> {
  return applyParagraphStyle({ name: list.name })
}
