// Inserting BwB blocks (callouts, key figures, tables, captions, the disclaimer) into the open document.

import type { Block } from './docxApi.ts'
import { buildBlockOoxml } from './docxApi.ts'
import { CAPTION_STYLES } from '../brand/styles.ts'
import { createIconRenderer } from './iconRenderer.ts'
import { isSeqField, updateFields } from './fields.ts'
import { runWord, supports } from './run.ts'

export type InsertAt = 'cursor' | 'end'

function usesIcons(blocks: Block[]): boolean {
  return JSON.stringify(blocks).includes('icon')
}

function hasCaptions(blocks: Block[]): boolean {
  return blocks.some((b) => (b.type === 'table' || b.type === 'figure') && Boolean(b.caption))
}

/** The numbers the next table and figure captions should show until Word refreshes the SEQ fields. */
async function nextCaptionNumbers(): Promise<{ Table: number; Figure: number }> {
  return runWord(async (context) => {
    const paragraphs = context.document.body.paragraphs
    paragraphs.load('items/style')
    await context.sync()
    let tables = 0
    let figures = 0
    for (const p of paragraphs.items) {
      if (p.style === CAPTION_STYLES.table.name) tables += 1
      else if (p.style === CAPTION_STYLES.figure.name) figures += 1
    }
    return { Table: tables + 1, Figure: figures + 1 }
  })
}

export interface InsertResult {
  warnings: string[]
}

/**
 * Inserts blocks below the paragraph holding the cursor (or in its place when it is empty), or at the end of
 * the document. Inside a table, blocks go after the table so they are never nested in a cell.
 */
export async function insertBlocks(blocks: Block[], at: InsertAt = 'cursor'): Promise<InsertResult> {
  const seqStart = hasCaptions(blocks) ? await nextCaptionNumbers() : undefined
  const icons = usesIcons(blocks) ? await createIconRenderer() : undefined
  const { ooxml, warnings } = await buildBlockOoxml(blocks, {
    ...(seqStart ? { seqStart } : {}),
    ...(icons ? { icons } : {}),
  })

  await runWord(async (context) => {
    let inserted: Word.Range
    if (at === 'end') {
      inserted = context.document.body.insertOoxml(ooxml, 'End')
    } else {
      const selection = context.document.getSelection()
      const paragraph = selection.paragraphs.getFirst()
      paragraph.load('text,tableNestingLevel')
      await context.sync()
      if (paragraph.tableNestingLevel > 0) {
        const after = paragraph.parentTable.insertParagraph('', 'After')
        inserted = after.insertOoxml(ooxml, 'Replace')
      } else if (paragraph.text.trim() === '') {
        inserted = paragraph.insertOoxml(ooxml, 'Replace')
      } else {
        inserted = paragraph.getRange('Whole').insertOoxml(ooxml, 'After')
      }
    }
    inserted.select('End')
    await context.sync()
  })

  if (seqStart && supports('1.5')) await updateFields(isSeqField)
  return { warnings }
}

/** Text at the cursor or in the selection, for forms that start from it. */
export async function selectedText(): Promise<string> {
  return runWord(async (context) => {
    const range = context.document.getSelection()
    range.load('text')
    await context.sync()
    return range.text
  })
}

/** Adds a real Word footnote at the cursor. Needs WordApi 1.5. */
export async function insertFootnote(text: string): Promise<void> {
  await runWord(async (context) => {
    context.document.getSelection().insertFootnote(text)
    await context.sync()
  })
}
