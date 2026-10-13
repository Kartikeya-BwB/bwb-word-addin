// Reads everything the Pre-QA check needs from the open document in a few batched calls.

import type { ControlInfo, DocSnapshot, HeaderFooterInfo, ParagraphInfo, PictureInfo, RunFormatting, TableInfo } from '../check/types.ts'
import { scanRunFormatting } from '../check/runFormatting.ts'
import { hasBwbStyles } from './styles.ts'
import { documentFileName, readDocInfo } from './properties.ts'
import { runWord, supports } from './run.ts'

const PARAGRAPH_PROPS = [
  'text', 'style', 'styleBuiltIn', 'isListItem', 'tableNestingLevel',
  'font/name', 'font/size', 'font/color', 'font/highlightColor', 'font/bold',
].map((p) => `items/${p}`).join(',')

function str(value: unknown): string {
  return typeof value === 'string' ? value : ''
}

function nullable<T>(value: T | null | undefined): T | null {
  return value === undefined || value === '' ? null : value
}

function parseXml(xml: string): Document {
  return new DOMParser().parseFromString(xml, 'application/xml')
}

async function countOptional(read: () => Promise<number>): Promise<number | null> {
  try {
    return await read()
  } catch {
    return null
  }
}

export async function takeSnapshot(options: { deepScan?: boolean } = {}): Promise<DocSnapshot> {
  const info = await readDocInfo()
  const bwbStyles = await hasBwbStyles().catch(() => null)

  const core = await runWord(async (context) => {
    const body = context.document.body
    const paragraphs = body.paragraphs
    paragraphs.load(PARAGRAPH_PROPS)
    const tables = body.tables
    tables.load('items/style,items/rowCount')
    const pictures = body.inlinePictures
    pictures.load('items/altTextTitle,items/altTextDescription')
    const controls = body.contentControls
    controls.load('items/tag,items/title,items/text')
    const sections = context.document.sections
    sections.load('items')
    await context.sync()

    const after = tables.items.map((t) => t.getParagraphAfterOrNullObject())
    const before = tables.items.map((t) => t.getParagraphBeforeOrNullObject())
    const firstCells = tables.items.map((t) => t.getCellOrNullObject(0, 0))
    after.forEach((p) => p.load('text,style'))
    before.forEach((p) => p.load('text,style'))
    firstCells.forEach((c) => c.load('value'))
    const headerFooters = sections.items.flatMap((section, i) =>
      (['header', 'footer'] as const).map((kind) => {
        const part = kind === 'header' ? section.getHeader('Primary') : section.getFooter('Primary')
        part.load('text')
        const pics = part.inlinePictures
        pics.load('items')
        return { section: i, kind, part, pics }
      }),
    )
    await context.sync()

    const paragraphInfo: ParagraphInfo[] = paragraphs.items.map((p, index) => ({
      index,
      text: p.text,
      style: str(p.style),
      styleBuiltIn: str(p.styleBuiltIn),
      isListItem: Boolean(p.isListItem),
      inTable: p.tableNestingLevel > 0,
      font: {
        name: nullable(p.font.name),
        size: typeof p.font.size === 'number' ? p.font.size : null,
        color: nullable(p.font.color),
        highlight: nullable(p.font.highlightColor),
        bold: typeof p.font.bold === 'boolean' ? p.font.bold : null,
      },
    }))
    const neighbour = (p: Word.Paragraph): { text: string; style: string } | null =>
      p.isNullObject ? null : { text: p.text, style: str(p.style) }
    const tableInfo: TableInfo[] = tables.items.map((t, index) => ({
      index,
      style: str(t.style),
      rows: t.rowCount,
      firstCell: firstCells[index] && !firstCells[index].isNullObject ? firstCells[index].value : '',
      after: after[index] ? neighbour(after[index]) : null,
      before: before[index] ? neighbour(before[index]) : null,
    }))
    const pictureInfo: PictureInfo[] = pictures.items.map((p, index) => ({
      index,
      altTitle: str(p.altTextTitle),
      altDescription: str(p.altTextDescription),
    }))
    const controlInfo: ControlInfo[] = controls.items.map((c, index) => ({ index, tag: str(c.tag), title: str(c.title), text: c.text }))
    const hf: HeaderFooterInfo[] = headerFooters.map((h) => ({
      section: h.section,
      kind: h.kind,
      text: h.part.text,
      pictures: h.pics.items.length,
    }))
    return { paragraphInfo, tableInfo, pictureInfo, controlInfo, hf }
  })

  const fieldCodes = supports('1.5')
    ? await runWord(async (context) => {
        const fields = context.document.body.fields
        fields.load('items/code')
        await context.sync()
        return fields.items.map((f) => f.code)
      }).catch(() => null)
    : null
  const comments = supports('1.4')
    ? await countOptional(() =>
        runWord(async (context) => {
          const items = context.document.body.getComments()
          items.load('items')
          await context.sync()
          return items.items.length
        }),
      )
    : null
  const trackedChanges = supports('1.6')
    ? await countOptional(() =>
        runWord(async (context) => {
          const items = context.document.body.getTrackedChanges()
          items.load('items')
          await context.sync()
          return items.items.length
        }),
      )
    : null

  let formatting: RunFormatting[] | null = null
  if (options.deepScan !== false) {
    try {
      const ooxml = await runWord(async (context) => {
        const result = context.document.body.getOoxml()
        await context.sync()
        return result.value
      })
      const scanned = scanRunFormatting(ooxml, parseXml)
      if (scanned) {
        const trusted = scanned.counted === core.paragraphInfo.length
        formatting = scanned.paragraphs.map((f) => (trusted ? f : { ...f, paragraph: null }))
      }
    } catch {
      formatting = null
    }
  }

  return {
    paragraphs: core.paragraphInfo,
    tables: core.tableInfo,
    pictures: core.pictureInfo,
    controls: core.controlInfo,
    headersFooters: core.hf,
    fieldCodes,
    comments,
    trackedChanges,
    formatting,
    info,
    fileName: documentFileName(),
    bwbStyles,
  }
}
