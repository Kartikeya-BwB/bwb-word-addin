// What each ribbon button does when it acts at once (ExecuteFunction). The names are the FunctionName values in
// ribbon.json; a test checks the two lists match. Each handler does the same as its task pane button.

import { defaultTableBlocks, findLibraryBlock } from '../brand/library.ts'
import type { ListStyle, ParagraphStyle } from '../brand/styles.ts'
import { HEADING_STYLES, LIST_STYLES, TEXT_STYLES, UNNUMBERED_HEADING_STYLES } from '../brand/styles.ts'
import type { Block } from '../docx/types.ts'
import { cleanSelectionWording } from '../word/clean.ts'
import { updateFields } from '../word/fields.ts'
import { insertBlocks } from '../word/insert.ts'
import { HARVEY_BALLS, insertHarveyBall, insertPageBreak } from '../word/marks.ts'
import { insertContents, insertDisclaimer, insertEuAcknowledgement } from '../word/parts.ts'
import { supports } from '../word/run.ts'
import type { Sticker } from '../word/stickers.ts'
import { setSticker } from '../word/stickers.ts'
import { addBwbStyles, applyListStyle, applyParagraphStyle, restyleBuiltInStyles } from '../word/styles.ts'
import { formatTableAtCursor } from '../word/tables.ts'

export type Handler = () => Promise<unknown>

function find<T extends { id: string }>(list: T[], id: string): T {
  const item = list.find((x) => x.id === id)
  if (!item) throw new Error(`Unknown BwB style ${id}`)
  return item
}

const paragraphStyle = (list: ParagraphStyle[], id: string): Handler => {
  const style = find(list, id)
  return () => applyParagraphStyle(style)
}

const listStyle = (id: string): Handler => {
  const style: ListStyle = find(LIST_STYLES, id)
  return () => applyListStyle(style)
}

const library = (id: string): Handler => {
  const block = findLibraryBlock(id)
  if (!block) throw new Error(`Unknown library block ${id}`)
  return () => insertBlocks(block.blocks())
}

const blocks = (make: () => Block[]): Handler => () => insertBlocks(make())

const harvey = (i: number): Handler => {
  const ball = HARVEY_BALLS[i]
  if (!ball) throw new Error(`No Harvey ball ${i}`)
  return () => insertHarveyBall(ball.glyph)
}

const sticker = (label: Sticker | null): Handler => () => setSticker(label)

export const HANDLERS: Record<string, Handler> = {
  addStyles: async () => {
    await addBwbStyles()
    if (supports('1.5')) await restyleBuiltInStyles()
  },

  styleHeading1: paragraphStyle(HEADING_STYLES, 'Heading1'),
  styleHeading2: paragraphStyle(HEADING_STYLES, 'Heading2'),
  styleHeading3: paragraphStyle(HEADING_STYLES, 'Heading3'),
  styleHeading4: paragraphStyle(HEADING_STYLES, 'Heading4'),
  styleHeading1Unnumbered: paragraphStyle(UNNUMBERED_HEADING_STYLES, 'BwBHeading1Unnumbered'),
  styleHeading2Unnumbered: paragraphStyle(UNNUMBERED_HEADING_STYLES, 'BwBHeading2Unnumbered'),
  styleNormal: paragraphStyle(TEXT_STYLES, 'Normal'),
  styleLead: paragraphStyle(TEXT_STYLES, 'BwBLead'),
  styleAccent: paragraphStyle(TEXT_STYLES, 'BwBAccent'),
  styleSource: paragraphStyle(TEXT_STYLES, 'BwBSource'),
  stylePullQuote: paragraphStyle(TEXT_STYLES, 'BwBPullQuote'),

  listGreenSquare: listStyle('BwBBulletGreenSquare'),
  listNavyCircle: listStyle('BwBBulletNavyCircle'),
  listGreenCircle: listStyle('BwBBulletGreenCircle'),
  listCharcoalDash: listStyle('BwBBulletCharcoalDash'),
  listArrow: listStyle('BwBBulletArrow'),
  listNumbered: listStyle('BwBNumberedNavy'),
  listLimeOnNavy: listStyle('BwBBulletLimeNavy'),

  cleanWording: cleanSelectionWording,

  blockCallout: library('callout'),
  blockNote: library('note'),
  blockStats: library('stats'),
  blockSnapshot: library('snapshot'),
  blockQuote: library('quote'),
  blockCase: library('case'),
  blockGrid: library('grid'),
  blockDetails: library('details'),

  tableInsert: blocks(defaultTableBlocks),
  tableFormat: () => formatTableAtCursor('standard'),
  captionTable: blocks(() => [{ type: 'caption', kind: 'Table', text: '{field:Title}', source: '{field:Source}' }]),
  captionFigure: blocks(() => [{ type: 'caption', kind: 'Figure', text: '{field:Title}', source: '{field:Source}' }]),
  sourceLine: blocks(() => [{ type: 'source', text: 'Source: {field:Source}' }]),

  harvey0: harvey(0),
  harvey1: harvey(1),
  harvey2: harvey(2),
  harvey3: harvey(3),
  harvey4: harvey(4),

  partContents: insertContents,
  partDisclaimer: insertDisclaimer,
  partEu: insertEuAcknowledgement,
  partPageBreak: insertPageBreak,

  stickerDraft: sticker('Draft'),
  stickerForDiscussion: sticker('For discussion'),
  stickerPreliminary: sticker('Preliminary'),
  stickerIllustrative: sticker('Illustrative'),
  stickerConfidential: sticker('Confidential'),
  stickerNotForDistribution: sticker('Not for distribution'),
  stickerRemove: sticker(null),

  updateAllFields: () => updateFields(),
}
