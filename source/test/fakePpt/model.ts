// The in-memory presentation behind the fake PowerPoint. Plain data, so tests build decks with the helpers below
// and inspect what the add-in did. See index.ts for how the pieces fit together.

export type FakeShapeType =
  | 'GeometricShape' | 'TextBox' | 'Image' | 'Graphic' | 'Placeholder' | 'Group' | 'Line' | 'Table' | 'Chart'
  | 'SmartArt' | 'Freeform' | 'Callout' | 'Media' | 'Unsupported'

export interface FakeFont {
  name: string | null
  size: number | null
  /** '#RRGGBB', upper case. */
  color: string | null
  bold: boolean | null
  italic: boolean | null
  underline: string | null
}

/** A stretch of text with one font. A text frame's text is its runs joined. */
export interface FakeRun {
  text: string
  font: FakeFont
}

export interface FakeTextFrame {
  runs: FakeRun[]
  /** One alignment for the whole frame: the fake does not keep paragraph formatting per paragraph. */
  horizontalAlignment: string
  /** Bullets shown or not, for the whole frame (paragraphFormat.bulletFormat.visible). */
  bulletVisible: boolean
  verticalAlignment: string
  autoSizeSetting: string
  wordWrap: boolean
  leftMargin: number
  rightMargin: number
  topMargin: number
  bottomMargin: number
}

export interface FakeFill {
  /** NoFill, Solid, Gradient, Pattern, PictureAndTexture or SlideBackground. */
  type: string
  /** '#RRGGBB', upper case. */
  foregroundColor: string
  transparency: number
}

export interface FakeLine {
  visible: boolean
  color: string
  weight: number
  dashStyle: string
  style: string
  transparency: number
}

export interface FakeTable {
  values: string[][]
  /** What addTable received: uniformCellProperties, specificCellProperties, style, columns, rows and so on. */
  options: Record<string, unknown>
  /** Formatting written through table cells (PowerPointApi 1.9), keyed "row,column". */
  cells: Record<string, Record<string, unknown>>
}

export interface FakeShape {
  id: string
  name: string
  type: FakeShapeType
  /** The unrotated frame, in points. For a group these are derived from its members: use boundsOf(). */
  left: number
  top: number
  width: number
  height: number
  /** Degrees clockwise, 0 to 360. */
  rotation: number
  visible: boolean
  /** Marked as decorative for screen readers (PowerPointApi 1.10). */
  isDecorative: boolean
  /** GeometricShape: Rectangle, Ellipse, Chevron and so on. */
  geometricShapeType?: string
  /** Line: Straight, Elbow or Curve. */
  connectorType?: string
  /** Placeholder: Title, CenterTitle, Subtitle, Body, Content and so on. */
  placeholderType?: string
  fill: FakeFill
  line: FakeLine
  /** null for shapes that hold no text: pictures, lines, groups and tables. */
  textFrame: FakeTextFrame | null
  /** Keys in upper case, as PowerPoint stores them. */
  tags: Map<string, string>
  /** Group: its members, back to front. */
  members?: FakeShape[]
  table?: FakeTable
  altTextTitle: string
  altTextDescription: string
}

export interface FakeLayout {
  id: string
  name: string
  type: string
  /** Copied onto a slide added with this layout. */
  placeholders: FakeShape[]
}

export interface FakeMaster {
  id: string
  name: string
  layouts: FakeLayout[]
}

export interface FakeSlide {
  id: string
  /** Top-level shapes, back to front: the index is the z-order position. */
  shapes: FakeShape[]
  tags: Map<string, string>
  layoutId: string
  masterId: string
}

export interface FakeSelection {
  /** The selected slides. The first is the active slide, the one in the editing pane. */
  slideIds: string[]
  /** The selected shapes on the active slide. */
  shapeIds: string[]
  /** A text selection, or the cursor, inside a shape. */
  textRange: { shapeId: string; start: number; length: number } | null
}

export interface FakeDataInsert {
  data: string
  /** 'image', 'xmlSvg', 'text' and so on (Office.CoercionType values). */
  coercionType: string
  options: Record<string, unknown>
  /** The shape the insert added, if any. */
  shapeId: string | null
}

export interface FakePresentation {
  id: string
  title: string
  slides: FakeSlide[]
  masters: FakeMaster[]
  tags: Map<string, string>
  pageSetup: { slideWidth: number; slideHeight: number }
  selection: FakeSelection
  /** The newest PowerPointApi requirement set this PowerPoint supports, such as '1.5' or '1.10'. */
  apiVersion: string
  /** Other requirement sets and their newest version, such as ImageCoercion '1.2'. */
  otherSets: Record<string, string>
  /** Every Office.context.document.setSelectedDataAsync call, in order. */
  dataInserts: FakeDataInsert[]
  /** Every insertSlidesFromBase64 call, in order. */
  slideInserts: Array<{ base64: string; options: Record<string, unknown> }>
  /** The slide ids of every exportAsBase64Presentation call, in order. */
  slideExports: string[][]
  /** Every PowerPoint.createPresentation call: its base64 file, or null when called without one. */
  createdPresentations: Array<string | null>
  /** Optional: turns an insertSlidesFromBase64 call into slides, so tests see them in the deck. */
  slidesFromBase64?: (base64: string, options: Record<string, unknown>) => FakeSlide[]
  /** Set it to make the next setSelectedDataAsync fail with this message. */
  failNextDataInsert: string | null
  settings: Map<string, unknown>
  handlers: Array<{ eventType: string; handler: (args: unknown) => unknown }>
}

let counter = 1000

/** A fresh id. Builders and the fake share one counter, so ids never collide within a test run. */
export function nextId(): string {
  counter += 1
  return String(counter)
}

export const font = (over: Partial<FakeFont> = {}): FakeFont => ({
  name: 'Aptos', size: 18, color: '#000000', bold: false, italic: false, underline: 'None', ...over,
})

export type FrameInit = Partial<Omit<FakeTextFrame, 'runs'>> & { font?: Partial<FakeFont>; runs?: FakeRun[] }

/** A text frame. Pass runs to start with mixed formatting. */
export function textFrame(text = '', over: FrameInit = {}): FakeTextFrame {
  const { font: f, runs, ...rest } = over
  return {
    runs: runs ?? [{ text, font: font(f) }],
    horizontalAlignment: 'Left', bulletVisible: false, verticalAlignment: 'Top', autoSizeSetting: 'AutoSizeNone', wordWrap: true,
    leftMargin: 7.2, rightMargin: 7.2, topMargin: 3.6, bottomMargin: 3.6,
    ...rest,
  }
}

export const run = (text: string, over: Partial<FakeFont> = {}): FakeRun => ({ text, font: font(over) })

export interface ShapeInit extends Partial<Omit<FakeShape, 'tags' | 'textFrame'>> {
  /** Text for a new text frame (ignored when textFrame is given). */
  text?: string
  font?: Partial<FakeFont>
  textFrame?: FakeTextFrame | null
  tags?: Record<string, string>
}

const solid = (color: string): FakeFill => ({ type: 'Solid', foregroundColor: color.toUpperCase(), transparency: 0 })
const noFill = (): FakeFill => ({ type: 'NoFill', foregroundColor: '#FFFFFF', transparency: 0 })
const outline = (visible: boolean, color = '#042433', weight = 1): FakeLine => ({
  visible, color, weight, dashStyle: 'Solid', style: 'Single', transparency: 0,
})

function tagMap(tags: Record<string, string> | undefined): Map<string, string> {
  return new Map(Object.entries(tags ?? {}).map(([k, v]) => [k.toUpperCase(), v]))
}

function build(defaults: Omit<FakeShape, 'id' | 'tags' | 'altTextTitle' | 'altTextDescription' | 'visible' | 'rotation' | 'isDecorative'>, init: ShapeInit): FakeShape {
  const { text, font: f, textFrame: frame, tags, ...rest } = init
  const id = rest.id ?? nextId()
  const shape: FakeShape = {
    ...defaults,
    id,
    rotation: 0,
    visible: true,
    isDecorative: false,
    altTextTitle: '',
    altTextDescription: '',
    ...rest,
    name: rest.name ?? `${defaults.name} ${id}`,
    tags: tagMap(tags),
  }
  if (frame !== undefined) shape.textFrame = frame
  else if (shape.textFrame && (text !== undefined || f)) {
    const before = shape.textFrame.runs
    shape.textFrame.runs = [{ text: text ?? before.map((r) => r.text).join(''), font: { ...(before[0]?.font ?? font()), ...f } }]
  }
  return shape
}

/** A rectangle (or another geometric shape) with PowerPoint's default style: accent fill, white centred text. */
export function shape(init: ShapeInit = {}): FakeShape {
  return build({
    name: init.geometricShapeType ?? 'Rectangle', type: 'GeometricShape', geometricShapeType: 'Rectangle',
    left: 0, top: 0, width: 100, height: 50, fill: solid('#156082'), line: outline(true),
    textFrame: textFrame('', { horizontalAlignment: 'Center', verticalAlignment: 'Middle', font: { color: '#FFFFFF' } }),
  }, init)
}

export function textBox(text: string, init: ShapeInit = {}): FakeShape {
  return build({
    name: 'TextBox', type: 'TextBox', left: 0, top: 0, width: 200, height: 30, fill: noFill(), line: outline(false),
    textFrame: textFrame(text, { autoSizeSetting: 'AutoSizeShapeToFitText' }),
  }, init)
}

/** A layout placeholder holding text: Title, CenterTitle, Subtitle, Body and so on. Body and content have bullets. */
export function placeholder(placeholderType: string, text = '', init: ShapeInit = {}): FakeShape {
  const title = /Title$/.test(placeholderType) && placeholderType !== 'Subtitle'
  const content = /Body|Content/.test(placeholderType)
  return build({
    name: title ? 'Title' : content ? 'Content Placeholder' : `${placeholderType} Placeholder`,
    type: 'Placeholder', placeholderType, left: 40, top: title ? 28 : 108, width: 880, height: title ? 64 : 360,
    fill: noFill(), line: outline(false),
    textFrame: textFrame(text, { bulletVisible: content, font: title ? { name: 'Candara', size: 24, bold: true } : { size: 14 } }),
  }, init)
}

/** A picture: type Image, as Office.js reports it (PowerPoint.ShapeType has no Picture). */
export function picture(init: ShapeInit = {}): FakeShape {
  return build({ name: 'Picture', type: 'Image', left: 0, top: 0, width: 100, height: 100, fill: noFill(), line: outline(false), textFrame: null }, init)
}

export function line(init: ShapeInit = {}): FakeShape {
  return build({
    name: 'Straight Connector', type: 'Line', connectorType: 'Straight', left: 0, top: 0, width: 100, height: 0,
    fill: noFill(), line: outline(true, '#156082', 0.75), textFrame: null,
  }, init)
}

export function group(members: FakeShape[], init: ShapeInit = {}): FakeShape {
  return build({ name: 'Group', type: 'Group', members, left: 0, top: 0, width: 0, height: 0, fill: noFill(), line: outline(false), textFrame: null }, init)
}

export function tableShape(values: string[][], init: ShapeInit = {}): FakeShape {
  return build({
    name: 'Table', type: 'Table', left: 0, top: 0, width: 100 * (values[0]?.length ?? 1), height: 30 * values.length,
    fill: noFill(), line: outline(false), textFrame: null, table: { values, options: {}, cells: {} },
  }, init)
}

export function slide(shapes: FakeShape[] = [], init: Partial<Omit<FakeSlide, 'tags'>> & { tags?: Record<string, string> } = {}): FakeSlide {
  const { tags, ...rest } = init
  return { id: `slide-${nextId()}`, shapes, layoutId: 'layout-title-content', masterId: 'master-1', ...rest, tags: tagMap(tags) }
}

/** The Office Theme master: Title Slide, Title and Content, Section Header, Title Only and Blank. */
export function defaultMasters(): FakeMaster[] {
  const layout = (id: string, name: string, type: string, placeholders: () => FakeShape[]): FakeLayout => ({ id, name, type, placeholders: placeholders() })
  return [{
    id: 'master-1', name: 'Office Theme',
    layouts: [
      layout('layout-title', 'Title Slide', 'Title', () => [placeholder('CenterTitle'), placeholder('Subtitle')]),
      layout('layout-title-content', 'Title and Content', 'Object', () => [placeholder('Title'), placeholder('Content')]),
      layout('layout-section', 'Section Header', 'SectionHeader', () => [placeholder('Title'), placeholder('Body')]),
      layout('layout-title-only', 'Title Only', 'TitleOnly', () => [placeholder('Title')]),
      layout('layout-blank', 'Blank', 'Blank', () => []),
    ],
  }]
}

/** A 16:9 deck (960 by 540 points), first slide active, nothing selected, every PowerPointApi set up to 1.10. */
export function createPresentation(slides: FakeSlide[] = [slide()], over: Partial<FakePresentation> = {}): FakePresentation {
  return {
    id: 'presentation-1',
    title: 'Deck',
    slides,
    masters: defaultMasters(),
    tags: new Map(),
    pageSetup: { slideWidth: 960, slideHeight: 540 },
    selection: { slideIds: slides[0] ? [slides[0].id] : [], shapeIds: [], textRange: null },
    apiVersion: '1.10',
    otherSets: { ImageCoercion: '1.2' },
    dataInserts: [],
    slideInserts: [],
    slideExports: [],
    createdPresentations: [],
    failNextDataInsert: null,
    settings: new Map(),
    handlers: [],
    ...over,
  }
}
