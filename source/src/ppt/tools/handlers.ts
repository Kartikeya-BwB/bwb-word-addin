// Ribbon handlers for the PowerPoint power tools that need no input from the pane. Kept apart from PPT_HANDLERS
// until their ribbon buttons exist, because every handler must have a button (test/commands.test.ts checks it).

import { applyFormat, pickUpFormat } from './painter.ts'
import { selectSimilar } from './similar.ts'
import { addTableTotalRow, equaliseColumnWidths, formatSelectedTable, transposeSelectedTable } from './tables.ts'

export type PptToolHandler = () => Promise<unknown>

export const PPT_TOOL_HANDLERS: Record<string, PptToolHandler> = {
  pptPainterPickUp: pickUpFormat,
  pptPainterApply: () => applyFormat(),
  pptTableFormat: formatSelectedTable,
  pptTableTranspose: transposeSelectedTable,
  pptTableTotalRow: addTableTotalRow,
  pptTableEqualWidths: equaliseColumnWidths,
  pptSelectSimilarFill: () => selectSimilar('fill'),
}
