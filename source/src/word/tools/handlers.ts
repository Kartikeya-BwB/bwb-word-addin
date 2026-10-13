// What the ribbon's Tools group buttons do (ExecuteFunction). The names are the FunctionName values in
// ribbon-word.json; merge these into HANDLERS in src/commands/handlers.ts. Each does the same as its pane button.

import { altTextFromCaptions, captionSelectedPicture } from './captions.ts'
import { clearStrayFormatting, joinBrokenLines } from './cleanup.ts'
import { alignPictures, matchPictureWidths, setPictureWidth } from './pictures.ts'
import { DEFAULT_NOTE, acceptAllChanges, addReviewNote, deleteAllComments, toggleTracking } from './review.ts'
import { tableToText, textToTable } from './tableData.ts'
import { addTotalRow, centreTable, deleteEmptyRows, distributeColumns, fitTableToWindow } from './tableTools.ts'
import { addNonBreakingSpaces, changeSelectionCase } from './textTools.ts'

export const WORD_TOOL_HANDLERS: Record<string, () => Promise<unknown>> = {
  wordPicturesFitWidth: () => setPictureWidth('text'),
  wordPicturesSameWidth: matchPictureWidths,
  wordPicturesCentre: () => alignPictures('Centered'),
  wordPicturesCaption: () => captionSelectedPicture(),
  wordPicturesAltFromCaptions: altTextFromCaptions,

  wordTableDistribute: distributeColumns,
  wordTableFitWindow: fitTableToWindow,
  wordTableCentre: centreTable,
  wordTableAddTotal: addTotalRow,
  wordTableDeleteEmptyRows: deleteEmptyRows,
  wordTextToTable: textToTable,
  wordTableToText: tableToText,

  wordCaseSentence: () => changeSelectionCase('sentence'),
  wordCaseTitle: () => changeSelectionCase('title'),
  wordCaseUpper: () => changeSelectionCase('upper'),
  wordCaseLower: () => changeSelectionCase('lower'),
  wordJoinBrokenLines: joinBrokenLines,
  wordClearFormatting: clearStrayFormatting,
  wordNonBreakingSpaces: addNonBreakingSpaces,

  wordToggleTracking: toggleTracking,
  wordAcceptAllChanges: acceptAllChanges,
  wordDeleteAllComments: deleteAllComments,
  wordAddReviewNote: () => addReviewNote(DEFAULT_NOTE),
}
