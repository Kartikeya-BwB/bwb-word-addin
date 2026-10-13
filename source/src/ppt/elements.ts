// BwB slide elements: callout, key figures, quote, process, matrix, timeline, step markers, source line and table.
// Each lands in the content area of the current slide, named "BwB ..." and tagged BWB_ELEMENT, ready to type over.

import { TAGS } from '../brand/slideGrid.ts'
import { insertElement, insertTable, taggedShapes, currentSlide } from './render.ts'
import { runPpt, supportsPpt } from './run.ts'
import {
  calloutSpec, keyFiguresSpec, matrixSpec, processSpec, quoteSpec, sourceLineSpec, stepMarkersSpec, tableSpec, timelineSpec,
} from './specs.ts'
import type { ElementSpec } from './specs.ts'

const TYPE_OVER = 'Type over the text.'

async function place(spec: ElementSpec, done: string): Promise<string> {
  const { groupId } = await insertElement(spec)
  const grouped = spec.group && spec.shapes.length > 1 && !groupId ? ' Your PowerPoint cannot group shapes from an add-in, so its parts are separate.' : ''
  return `${done} ${TYPE_OVER}${grouped}`
}

export const insertCallout = (): Promise<string> => place(calloutSpec(), 'Callout added at the foot of the content area.')
export const insertKeyFigures = (): Promise<string> => place(keyFiguresSpec(), 'Key figures added.')
export const insertQuote = (): Promise<string> => place(quoteSpec(), 'Quote added.')
export const insertProcess = (): Promise<string> => place(processSpec(), 'Process added.')
export const insertMatrix = (): Promise<string> => place(matrixSpec(supportsPpt('1.10')), 'Matrix added.')
export const insertTimeline = (): Promise<string> => place(timelineSpec(), 'Timeline added.')

export async function insertStepMarkers(): Promise<string> {
  await insertElement(stepMarkersSpec())
  return 'Numbered markers added. Drag each one to its step.'
}

/** One source line per slide: when there is one already, it is selected instead. */
export async function insertSourceLine(): Promise<string> {
  const existing = await runPpt(async (context) => {
    const slide = await currentSlide(context)
    const found = (await taggedShapes(context, [slide], TAGS.element)).find((t) => t.value === 'source')
    if (!found) return false
    found.shape.load('id')
    await context.sync()
    slide.setSelectedShapes([found.shape.id])
    await context.sync()
    return true
  })
  if (existing) return 'This slide has a source line already, so it is selected.'
  await insertElement(sourceLineSpec())
  return 'Source line added under the content. Type over the text.'
}

export async function insertBwbTable(): Promise<string> {
  await insertTable(tableSpec())
  return `Table added. ${supportsPpt('1.9') ? '' : 'Your PowerPoint keeps its own table style underneath; clear it under Table Design if it shows. '}Type in the cells.`
}

export interface ElementEntry {
  id: string
  label: string
  hint: string
  /** The PowerPointApi set it needs. */
  version: string
  insert: () => Promise<string>
}

export const ELEMENTS: ElementEntry[] = [
  { id: 'callout', label: 'Callout', hint: 'An off-white box with a green bar, a title and a message', version: '1.5', insert: insertCallout },
  { id: 'key-figures', label: 'Key figures', hint: 'Three headline numbers on navy, the first in lime', version: '1.5', insert: insertKeyFigures },
  { id: 'quote', label: 'Quote', hint: 'A quotation in Candara italic with a lime bar', version: '1.5', insert: insertQuote },
  { id: 'process', label: 'Process, three steps', hint: 'Chevrons across the content area', version: '1.5', insert: insertProcess },
  { id: 'matrix', label: 'Two-by-two matrix', hint: 'Four quadrants with axis labels', version: '1.5', insert: insertMatrix },
  { id: 'timeline', label: 'Timeline', hint: 'A line with four milestones', version: '1.5', insert: insertTimeline },
  { id: 'step-marker', label: 'Numbered markers', hint: 'Navy circles numbered 1 to 4', version: '1.5', insert: insertStepMarkers },
  { id: 'source', label: 'Source line', hint: 'Aptos 8 pt charcoal under the content', version: '1.5', insert: insertSourceLine },
  { id: 'table', label: 'BwB table', hint: 'Four by four, navy header row, light grey rules', version: '1.8', insert: insertBwbTable },
]
