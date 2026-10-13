/** Composes a whole BwB document from content (build_document.compose) and writes the .docx (buildDocx). */
import { Doc, fmtDate } from './builder.ts'
import { bullets, caption, figure, guidance, heading, numbered, paragraph, pullQuote, toInt } from './blocksText.ts'
import { callout, caseStudy, iconGrid, snapshot, stats } from './blocksPanels.ts'
import { keyValues, table } from './blocksTables.ts'
import { finishDocx } from './finish.ts'
import { Package } from './package.ts'
import { cover, disclaimer, DISCLAIMER_NOTE, divider, titleBlock, toc, tocBlock } from './sections.ts'
import type { Block, BuildAssets, BuildResult, Content, ContentMeta, KeyValueRows } from './types.ts'
import { pySplit, pyStrip } from './xml.ts'

// headings that stay unnumbered (front and back matter) when numbering is "auto"; the lookahead is Python's \b
const UNNUMBERED = new RegExp(
  '^(table of contents|contents|executive summary|summary|key messages|foreword|preface|acknowledg(e)?ments?|' +
    'abbreviations|acronyms|glossary|references|bibliography|sources|appendix|appendices|annex(es)?|about (this|the) ' +
    '(document|report|paper|note|template)|document control|version history|disclaimer|important notice)(?![\\p{L}\\p{N}_])',
  'iu',
)

export function isFrontMatter(headingText: string): boolean {
  return UNNUMBERED.test(pyStrip(headingText))
}

export function countHeadings(blocks: readonly Block[], maxLevel = 2): number {
  return blocks.filter((b) => b.type === 'heading' && toInt(b.level ?? 1, 1) <= maxLevel).length
}

function wordCount(blocks: readonly Block[]): number {
  let n = 0
  for (const b of blocks) {
    const t: unknown = (b as { text?: unknown }).text
    if (typeof t === 'string') n += pySplit(t).length
  }
  return n
}

function hasRows(rows: KeyValueRows | undefined): rows is KeyValueRows {
  return rows !== undefined && (Array.isArray(rows) ? rows.length > 0 : Object.keys(rows).length > 0)
}

/** Add one content block (Doc.add_block), including the add-in's extension blocks. */
export function addBlock(doc: Doc, block: Block): void {
  const type: unknown = (block as { type?: unknown }).type
  const b = (type === undefined ? { ...block, type: 'paragraph' } : block) as Block
  switch (b.type) {
    case 'heading':
      return heading(doc, b.text, b.level ?? 1, b.numbered ?? true, b.in_toc ?? true)
    case 'paragraph':
      return paragraph(doc, b.text)
    case 'lead':
      return paragraph(doc, b.text, 'BwBLead')
    case 'source':
      return paragraph(doc, b.text, 'BwBSource')
    case 'accent':
      return paragraph(doc, b.text, 'BwBAccent')
    case 'bullets':
      return bullets(doc, b.items ?? [], b.style ?? 'green-square')
    case 'numbered':
      return numbered(doc, b.items ?? [])
    case 'table':
      return table(doc, b)
    case 'figure':
      if (!(b.data instanceof Uint8Array) || !b.data.length) return doc.warn('A figure without image data was skipped.')
      return figure(doc, b)
    case 'callout':
      return callout(doc, b.title, b.text ?? '', b.variant ?? 'green', b.bullets, b.icon)
    case 'pull_quote':
      return pullQuote(doc, b.text, b.attribution)
    case 'stats':
      return stats(doc, b.items ?? [], b.hero)
    case 'snapshot':
      return snapshot(doc, b.items ?? [], b.columns ?? 3)
    case 'case_study':
      return caseStudy(doc, b.title, b.meta, b.body ?? [], b.stats, b.source, b.icon)
    case 'icon_grid':
      return iconGrid(doc, b.items, b.columns ?? 3, b.variant ?? 'light', b.hero)
    case 'key_values':
      return keyValues(doc, b.rows ?? [])
    case 'divider':
      if (doc.fragment) return doc.warn('A section divider cannot be inserted into an open document; it was skipped.')
      return divider(doc, b.number ?? '', b.title, b.subtitle ?? '', b.variant ?? 'dark')
    case 'page_break':
      return doc.pageBreak()
    case 'guidance':
      return guidance(doc, b.text)
    case 'disclaimer':
      return disclaimer(doc)
    case 'toc':
      return tocBlock(doc, b.levels, b.title ?? 'Table of Contents')
    case 'caption':
      if (b.kind !== 'Table' && b.kind !== 'Figure') return doc.warn(`Caption kind '${String(b.kind)}' is not Table or Figure; skipped.`)
      return doc.add(caption(doc, b.kind, b.text, b.source))
    default:
      return doc.warn(`Unknown block type '${String((b as { type?: unknown }).type)}' skipped.`)
  }
}

/** Front matter, body and closing matter on a Doc (build_document.compose). */
export function compose(doc: Doc, content: Content): void {
  const meta: ContentMeta = content.meta ?? {}
  const blocks: readonly Block[] = content.blocks ?? []
  const date = meta.date || fmtDate(doc.now)
  const prepared = meta.prepared_by ?? 'Bankers without Boundaries'
  const metaLeft = meta.meta_left || `${prepared}  \u{B7}  ${date}`
  const title = meta.title || 'Untitled document'

  // ---- front matter
  if ((meta.format ?? 'report') === 'memo') {
    titleBlock(doc, title, meta.subtitle || meta.tagline || null, metaLeft)
    if (hasRows(meta.details)) keyValues(doc, meta.details)
  } else {
    cover(doc, title, meta.subtitle ?? '', {
      tagline: meta.tagline ?? null,
      metaLeft,
      metaRight: meta.meta_right ?? 'bwb.earth',
      image: meta.cover_image ?? null,
    })
    if (hasRows(meta.details)) {
      keyValues(doc, meta.details)
      doc.sectionStart = true
    }
    const tocOpt = meta.toc ?? 'auto'
    if (tocOpt === true || (tocOpt === 'auto' && countHeadings(blocks, 2) >= 5)) {
      toc(doc, toInt(meta.toc_levels ?? 2, 2), meta.toc_title ?? 'Table of Contents')
      doc.pageBreak() // the text starts on a fresh page after the contents
    }
  }

  // ---- body
  const numbering = meta.numbering ?? 'auto'
  const dividers = meta.dividers ?? 'none'
  const h1Opt = meta.h1_page_break ?? 'auto'
  doc.h1PageBreak = h1Opt === true ? true : h1Opt === false ? false : wordCount(blocks) >= 3000 || dividers === 'h1'
  let parentNumbered: unknown = true
  let h1Count = 0
  for (const block of blocks) {
    let b = block
    if (b.type === 'heading') {
      const lvl = toInt(b.level ?? 1, 1)
      if (!('numbered' in b)) {
        let n: boolean
        if (numbering === 'none') n = false
        else if (numbering === 'all') n = true
        else if (lvl === 1) n = !isFrontMatter(b.text)
        else n = Boolean(parentNumbered)
        b = { ...b, numbered: n }
      }
      if (lvl === 1) {
        parentNumbered = b.numbered
        if (dividers === 'h1' && b.numbered) {
          h1Count += 1
          divider(doc, String(h1Count).padStart(2, '0'), b.text, b.subtitle ?? '', b.divider_variant ?? 'dark')
        }
      }
    }
    addBlock(doc, b)
  }
  if (!blocks.length) doc.warn('The content has no blocks.')

  // ---- closing matter
  if (meta.disclaimer ?? true) {
    const paras = meta.disclaimer_paragraphs?.length ? meta.disclaimer_paragraphs : null
    disclaimer(doc, paras, meta.disclaimer_note ?? DISCLAIMER_NOTE)
  }
}

/**
 * A complete BwB .docx built on the base template (public/templates/bwb-base.docx). Reports take their cover photo from
 * assets.cover unless meta.cover_image is given. Word refreshes the contents and page numbers when it opens the file
 * (settings.xml asks it to update fields).
 */
export async function buildDocx(content: Content, assets: BuildAssets): Promise<BuildResult> {
  const now = assets.now ?? new Date()
  const extraMedia = new Map<string, Uint8Array>()
  if (assets.cover) extraMedia.set('cover_default.jpg', assets.cover)
  const doc = new Doc(Package.open(assets.base), { icons: assets.icons, extraMedia, now })
  compose(doc, content)
  const meta = content.meta ?? {}
  const docx = await finishDocx(
    doc,
    { title: meta.title ?? '', subject: meta.subject ?? '', keywords: meta.keywords ?? '', description: meta.description ?? '' },
    meta.custom_properties,
  )
  return { docx, warnings: [...doc.warnings] }
}
