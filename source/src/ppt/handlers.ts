// What each PowerPoint ribbon button does when it acts at once (ExecuteFunction). The names are the "fn" values in
// src/commands/ribbon-powerpoint.json; the agenda, tracker and section buttons live with the slide tools.
// Defaults: align lines up with the selection (one shape with the margins), distribute spaces out within the
// selection, same size takes the anchor's size, and the "Align to slide" items do what they say.

import type { AlignEdge, Reference } from '../geometry/index.ts'
import type { Sticker } from '../word/stickers.ts'
import {
  alignShapes, distributeShapes, groupShapes, matchShapeSize, orderShapes, snapShapes, stackShapes, swapShapes, tidyShapes, ungroupShapes,
} from './arrange.ts'
import {
  insertBwbTable, insertCallout, insertKeyFigures, insertMatrix, insertProcess, insertQuote, insertSourceLine, insertStepMarkers, insertTimeline,
} from './elements.ts'
import { applyBwbFonts, applyFill, applyOutline, applyTextColour, cleanSelectedText, setAutofitOff, setTextMargins } from './format.ts'
import { addReviewNote, insertHarveyBall, insertRagDot, insertStatusMark, removeAllNotes, removeAllStickers, setSlideStickers } from './marks.ts'

export type PptHandler = () => Promise<unknown>

const aligned = (edge: AlignEdge, reference: Reference = 'selection', asGroup = false): PptHandler => () => alignShapes(edge, reference, asGroup)
const sticker = (label: Sticker): PptHandler => () => setSlideStickers(label)

export const PPT_HANDLERS: Record<string, PptHandler> = {
  pptStickerDraft: sticker('Draft'),
  pptStickerForDiscussion: sticker('For discussion'),
  pptStickerPreliminary: sticker('Preliminary'),
  pptStickerIllustrative: sticker('Illustrative'),
  pptStickerConfidential: sticker('Confidential'),
  pptStickerNotForDistribution: sticker('Not for distribution'),
  pptStickerRemoveAll: removeAllStickers,

  pptAlignLeft: aligned('left'),
  pptAlignCentre: aligned('centre'),
  pptAlignRight: aligned('right'),
  pptAlignTop: aligned('top'),
  pptAlignMiddle: aligned('middle'),
  pptAlignBottom: aligned('bottom'),

  pptCentreOnSlide: aligned('centre', 'slide', true),
  pptMiddleOnSlide: aligned('middle', 'slide', true),
  pptAlignMarginLeft: aligned('left', 'margins', true),
  pptAlignMarginRight: aligned('right', 'margins', true),
  pptAlignMarginTop: aligned('top', 'content', true),
  pptAlignMarginBottom: aligned('bottom', 'content', true),

  pptDistributeHorizontally: () => distributeShapes({ axis: 'horizontal', within: 'selection' }),
  pptDistributeVertically: () => distributeShapes({ axis: 'vertical', within: 'selection' }),
  pptDistributeAcrossMargins: () => distributeShapes({ axis: 'horizontal', within: 'margins' }),
  pptDistributeDownContent: () => distributeShapes({ axis: 'vertical', within: 'content' }),
  pptTidyUp: tidyShapes,

  pptSameWidth: () => matchShapeSize('width', 'anchor'),
  pptSameHeight: () => matchShapeSize('height', 'anchor'),
  pptSameSize: () => matchShapeSize('both', 'anchor'),

  pptSwap: swapShapes,
  pptStackHorizontally: () => stackShapes('horizontal'),
  pptStackVertically: () => stackShapes('vertical'),
  pptSnapToGrid: snapShapes,
  pptBringToFront: () => orderShapes('BringToFront'),
  pptSendToBack: () => orderShapes('SendToBack'),
  pptGroup: groupShapes,
  pptUngroup: ungroupShapes,

  pptInsertCallout: insertCallout,
  pptInsertKeyFigures: insertKeyFigures,
  pptInsertQuote: insertQuote,
  pptInsertProcess: insertProcess,
  pptInsertMatrix: insertMatrix,
  pptInsertTimeline: insertTimeline,
  pptInsertStepMarkers: insertStepMarkers,
  pptInsertSourceLine: insertSourceLine,
  pptInsertTable: insertBwbTable,

  pptHarvey0: () => insertHarveyBall(0),
  pptHarvey1: () => insertHarveyBall(1),
  pptHarvey2: () => insertHarveyBall(2),
  pptHarvey3: () => insertHarveyBall(3),
  pptHarvey4: () => insertHarveyBall(4),

  pptStatusOnTrack: () => insertStatusMark('onTrack'),
  pptStatusAtRisk: () => insertStatusMark('atRisk'),
  pptStatusOffTrack: () => insertStatusMark('offTrack'),
  pptRagGreen: () => insertRagDot('green'),
  pptRagAmber: () => insertRagDot('amber'),
  pptRagRed: () => insertRagDot('red'),

  pptNoteAdd: () => addReviewNote(),
  pptNotesRemoveAll: removeAllNotes,
  pptCleanText: cleanSelectedText,

  pptFillNavy: () => applyFill('navy'),
  pptFillGreen: () => applyFill('green'),
  pptFillOffWhite: () => applyFill('offWhite'),
  pptFillNone: () => applyFill('none'),
  pptTextNavy: () => applyTextColour('navy'),
  pptTextCharcoal: () => applyTextColour('charcoal'),
  pptTextWhite: () => applyTextColour('white'),
  pptLineGrey: () => applyOutline('grey'),
  pptLineNone: () => applyOutline('none'),

  pptFontsBwB: applyBwbFonts,
  pptMarginsZero: () => setTextMargins('zero'),
  pptMarginsStandard: () => setTextMargins('standard'),
  pptAutofitOff: setAutofitOff,
}
