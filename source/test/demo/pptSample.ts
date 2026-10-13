// A sample BwB deck for the PowerPoint demo, with the kinds of slips the deck check and the Arrange tab are for:
// boxes that do not line up, an off-brand font, a title in a text box, US spelling, a sticker and a review note.

import type { FakeFill, FakeLine, FakePresentation, FakeShape, FakeSlide } from '../fakePpt/index.ts'
import { createPresentation, line, placeholder, run, shape, slide, tableShape, textBox, textFrame } from '../fakePpt/index.ts'

const NAVY = '#022338'
const GREEN = '#328840'
const LIME = '#BED62F'
const CHARCOAL = '#4A4A4C'
const OFF_WHITE = '#F7F8F5'
const WHITE = '#FFFFFF'

const solid = (color: string): FakeFill => ({ type: 'Solid', foregroundColor: color, transparency: 0 })
const noLine = (): FakeLine => ({ visible: false, color: NAVY, weight: 0, dashStyle: 'Solid', style: 'Single', transparency: 0 })

/** The green, charcoal and lime strip along the foot of every BwB slide. */
function strip(): FakeShape[] {
  const bar = (left: number, width: number, color: string): FakeShape =>
    shape({ name: 'Brand strip', left, top: 536, width, height: 4, fill: solid(color), line: noLine(), text: '' })
  return [bar(0, 480, GREEN), bar(480, 317, CHARCOAL), bar(797, 163, LIME)]
}

function footer(n: number): FakeShape[] {
  return [
    textBox('Bankers without Boundaries', { name: 'Footer logo', left: 40, top: 510, width: 200, height: 18, font: { size: 9, color: NAVY, bold: true } }),
    textBox(String(n), {
      name: 'Slide number', left: 880, top: 510, width: 40, height: 18,
      textFrame: textFrame(String(n), { horizontalAlignment: 'Right', font: { size: 9, color: CHARCOAL } }),
    }),
    ...strip(),
  ]
}

function titleSlide(): FakeSlide {
  return slide([
    shape({ name: 'Background', left: 0, top: 0, width: 960, height: 540, fill: solid(NAVY), line: noLine(), text: '' }),
    placeholder('CenterTitle', 'Scaling adaptation finance in East Africa', {
      left: 60, top: 170, width: 760, height: 120, font: { name: 'Candara', size: 40, bold: true, color: WHITE },
    }),
    shape({ name: 'Lime rule', left: 60, top: 300, width: 120, height: 4, fill: solid(LIME), line: noLine(), text: '' }),
    placeholder('Subtitle', 'Investment case for the steering group · 13 October 2026', {
      left: 60, top: 316, width: 760, height: 40, font: { name: 'Aptos', size: 16, color: WHITE },
    }),
  ], { id: 'slide-title', layoutId: 'layout-title' })
}

function optionBox(left: number, top: number, width: number, height: number, title: string, body: string, fill: string): FakeShape {
  return shape({
    name: title, geometricShapeType: 'Rectangle', left, top, width, height, fill: solid(fill), line: noLine(),
    textFrame: textFrame('', {
      runs: [run(`${title}\n`, { name: 'Aptos', size: 16, bold: true, color: WHITE }), run(body, { name: 'Aptos', size: 13, color: WHITE })],
      horizontalAlignment: 'Left', verticalAlignment: 'Top', leftMargin: 14, rightMargin: 14, topMargin: 12, bottomMargin: 12,
    }),
  })
}

function optionsSlide(): FakeSlide {
  return slide([
    placeholder('Title', 'Three routes could unlock USD 400m for adaptation by 2030', { font: { name: 'Candara', size: 24, bold: true, color: NAVY } }),
    optionBox(40, 120, 270, 190, 'Guarantee facility', 'Covers first losses so local banks lend to smallholders and water utilities.', NAVY),
    optionBox(338, 131, 250, 176, 'Resilience bond', 'A listed bond for municipal drainage and flood defences, with an outcome coupon.', GREEN),
    optionBox(622, 114, 292, 201, 'Insurance pool', 'Parametric cover for drought, paid out within 30 days of a trigger.', CHARCOAL),
    line({ name: 'Divider', left: 40, top: 350, width: 880, height: 0, line: { visible: true, color: '#E8E9E6', weight: 1, dashStyle: 'Solid', style: 'Single', transparency: 0 } }),
    textBox('Each route can start in 2027 and reach scale by 2030. The guarantee facility is the quickest to set up.', {
      name: 'Takeaway', left: 40, top: 366, width: 880, height: 50, font: { name: 'Aptos', size: 14, color: '#000000' },
    }),
    textBox('Source: BwB analysis of OECD and AfDB data, 2025', { name: 'Source', left: 40, top: 476, width: 700, height: 20, font: { name: 'Arial', size: 9, color: CHARCOAL } }),
    ...footer(2),
  ], { id: 'slide-options' })
}

function termsSlide(): FakeSlide {
  return slide([
    textBox('Organization of the facility \u2014 key terms', { name: 'Heading', left: 44, top: 34, width: 860, height: 50, font: { name: 'Calibri', size: 28, bold: true, color: NAVY } }),
    tableShape([
      ['Term', 'Guarantee facility', 'Resilience bond', 'Insurance pool'],
      ['Size', 'USD 150m', 'USD 200m', 'USD 50m'],
      ['First loss', '20%', 'None', '10%'],
      ['Tenor', '7 years', '12 years', '1 year, renewable'],
      ['Launch', '2027', '2028', '2027'],
    ], { name: 'Key terms', left: 40, top: 112, width: 880, height: 220 }),
    shape({
      name: 'Draft sticker', left: 820, top: 8, width: 100, height: 22, fill: solid(WHITE),
      line: { visible: true, color: NAVY, weight: 1, dashStyle: 'Solid', style: 'Single', transparency: 0 },
      text: 'Draft', font: { name: 'Aptos', size: 10, bold: true, color: NAVY }, tags: { BWB_STICKER: 'Draft', BWB_ELEMENT: 'sticker' },
    }),
    textBox('Source: BwB term sheet, draft of 6 October 2026', { name: 'Source', left: 40, top: 476, width: 700, height: 20, font: { name: 'Aptos', size: 8, color: CHARCOAL } }),
    ...footer(3),
  ], { id: 'slide-terms', layoutId: 'layout-title-only' })
}

function nextStepsSlide(): FakeSlide {
  return slide([
    placeholder('Title', 'Next steps: agree the term sheet by 30 November', { font: { name: 'Candara', size: 24, bold: true, color: NAVY } }),
    placeholder('Content', 'Confirm the two anchor investors\nAgree the first-loss share with the donors\nDraft the term sheet and the investment memo', {
      top: 108, height: 220, font: { name: 'Aptos', size: 16, color: '#000000' },
    }),
    shape({
      name: 'Review note', left: 520, top: 360, width: 400, height: 56, fill: solid(NAVY), line: noLine(),
      textFrame: textFrame('', {
        runs: [run('Note (13 October 2026): ', { name: 'Aptos', size: 11, bold: true, color: LIME }), run('check the dates with the client before Friday.', { name: 'Aptos', size: 11, color: WHITE })],
        horizontalAlignment: 'Left', verticalAlignment: 'Middle', leftMargin: 10, rightMargin: 10,
      }),
      tags: { BWB_NOTE: '2026-10-13', BWB_ELEMENT: 'note' },
    }),
    shape({ name: 'Callout', left: 40, top: 360, width: 440, height: 56, fill: solid(OFF_WHITE), line: noLine(), text: 'The donors meet on 18 November.', font: { name: 'Aptos', size: 13, color: NAVY } }),
    ...footer(4),
  ], { id: 'slide-next' })
}

export function samplePresentation(): FakePresentation {
  const deck = createPresentation([titleSlide(), optionsSlide(), termsSlide(), nextStepsSlide()], { title: 'Adaptation finance, steering group' })
  // Open on the slide with the boxes that need lining up, with all three selected.
  deck.selection = { slideIds: ['slide-options'], shapeIds: [], textRange: null }
  return deck
}
