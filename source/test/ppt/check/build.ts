// Builds deck snapshots for the rule tests: a BwB-clean shape by default, so each test changes only what it checks.

import { CONTENT_AREA, SOURCE_AREA, TITLE_AREA } from '../../../src/brand/slideGrid.ts'
import type { DeckSnapshot, FontInfo, ShapeInfo, SlideInfo } from '../../../src/ppt/check/types.ts'

export const font = (over: Partial<FontInfo> = {}): FontInfo => ({ name: 'Aptos', size: 14, colour: '000000', bold: false, italic: false, ...over })

let next = 1

export function shape(over: Partial<ShapeInfo> = {}): ShapeInfo {
  const id = over.id ?? String(next++)
  return {
    path: [id], level: 0, name: `TextBox ${id}`, type: 'TextBox', placeholder: null, contains: null,
    left: CONTENT_AREA.left, top: CONTENT_AREA.top, width: 400, height: 100,
    tags: {}, text: '', font: font(), pieces: [],
    fill: { type: 'NoFill', colour: null }, line: { visible: false, colour: null }, alt: null,
    ...over, id,
  }
}

export function text(value: string, over: Partial<ShapeInfo> = {}): ShapeInfo {
  return shape({ text: value, ...over })
}

export function title(value: string, over: Partial<ShapeInfo> = {}): ShapeInfo {
  return shape({
    type: 'Placeholder', placeholder: 'Title', name: 'Title 1', ...TITLE_AREA, text: value,
    font: font({ name: 'Candara', size: 24, bold: true, colour: '022338' }), ...over,
  })
}

export function body(value: string, over: Partial<ShapeInfo> = {}): ShapeInfo {
  return shape({ type: 'Placeholder', placeholder: 'Body', name: 'Content Placeholder 2', ...CONTENT_AREA, text: value, ...over })
}

export function source(value: string, over: Partial<ShapeInfo> = {}): ShapeInfo {
  return shape({ name: 'Source', ...SOURCE_AREA, text: value, font: font({ size: 8, colour: '4A4A4C' }), ...over })
}

export function picture(over: Partial<ShapeInfo> = {}): ShapeInfo {
  return shape({ type: 'Image', name: 'Picture 3', text: null, font: null, fill: null, line: null, alt: { title: '', description: '', decorative: false }, ...over })
}

export function slide(shapes: ShapeInfo[], over: Partial<SlideInfo> = {}): SlideInfo {
  return { id: '', index: 0, tags: {}, layout: 'Title and Content', shapes, ...over }
}

/** A deck from slides, numbering them and giving each an id like PowerPoint's ("256#1000"). */
export function deck(slides: SlideInfo[], over: Partial<DeckSnapshot> = {}): DeckSnapshot {
  return {
    slides: slides.map((s, index) => ({ ...s, index, id: s.id || `${256 + index}#${1000 + index}` })),
    width: 960, height: 540, selected: [],
    can: { formatting: true, placeholders: true, groups: true, altText: true, pageSize: true, selection: true },
    partial: false,
    ...over,
  }
}
