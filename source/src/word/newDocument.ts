// Starting a new BwB document: build the .docx in the browser from the BwB template and a document-type
// structure, then let Word open it as a new, unsaved document.

import type { DocType } from '../brand/docTypes.ts'
import type { ImageEntry } from '../brand/images.ts'
import { euAcknowledgementBlocks } from '../brand/eu.ts'
import { toBase64 } from '../media/base64.ts'
import type { Block, Content } from './docxApi.ts'
import { buildDocx } from './docxApi.ts'
import { createIconRenderer } from './iconRenderer.ts'
import type { FunderFrame } from './properties.ts'
import { runWord } from './run.ts'

export interface NewDocumentRequest {
  type: DocType
  title: string
  tagline: string
  date: string
  preparedBy: string
  /** Memo details table: To, From, Subject and so on. */
  details: Array<[string, string]>
  disclaimer: boolean
  /** Reports only: a navy divider page before each Heading 1. */
  dividers?: boolean
  /** Reports only: a photo from the BwB image library for the cover, instead of the default harbour photo. */
  coverPhoto?: ImageEntry
  project: string
  deliverable: string
  funder: FunderFrame | ''
}

async function fetchBytes(url: string): Promise<Uint8Array> {
  const response = await fetch(url)
  if (!response.ok) throw new Error(`Could not load ${url} (${response.status}).`)
  return new Uint8Array(await response.arrayBuffer())
}

export { toBase64 }

export function contentFor(request: NewDocumentRequest): Content {
  const { type } = request
  // EU deliverables carry the funding acknowledgement up front, before the executive summary.
  const blocks: Block[] = [...(request.funder === 'eu-grant' ? euAcknowledgementBlocks() : []), ...type.skeleton()]
  const custom: Record<string, string> = { 'BwB document type': type.id }
  if (request.project) custom['BwB project'] = request.project
  if (request.deliverable) custom['BwB deliverable'] = request.deliverable
  if (request.funder) custom['BwB funder'] = request.funder
  return {
    meta: {
      title: request.title.trim() || type.name,
      ...(request.tagline.trim() ? { tagline: request.tagline.trim() } : {}),
      format: type.layout,
      date: request.date,
      prepared_by: request.preparedBy.trim() || 'Bankers without Boundaries',
      ...(request.details.length ? { details: request.details } : {}),
      toc: 'auto',
      ...(request.dividers ? { dividers: 'h1' as const } : {}),
      disclaimer: request.disclaimer,
      subject: type.name,
      custom_properties: custom,
    },
    blocks,
  }
}

export async function createNewDocument(request: NewDocumentRequest): Promise<{ warnings: string[] }> {
  const report = request.type.layout === 'report'
  const photo = report ? request.coverPhoto : undefined
  const [base, cover, icons] = await Promise.all([
    fetchBytes('./templates/bwb-base.docx'),
    report ? fetchBytes(photo ? `./${photo.file}` : './templates/bwb-cover.jpg') : Promise.resolve(undefined),
    createIconRenderer().catch(() => undefined),
  ])
  const content = contentFor(request)
  if (photo && cover && content.meta) content.meta.cover_image = { data: cover, ext: 'jpg', widthPx: photo.width, heightPx: photo.height }
  const { docx, warnings } = await buildDocx(content, {
    base,
    ...(cover ? { cover } : {}),
    ...(icons ? { icons } : {}),
  })
  const base64 = toBase64(docx)
  await runWord(async (context) => {
    context.application.createDocument(base64).open()
    await context.sync()
  })
  return { warnings }
}

/** Today as BwB writes dates: 5 October 2026. */
export function bwbDate(date: Date = new Date()): string {
  return date.toLocaleDateString('en-GB', { day: 'numeric', month: 'long', year: 'numeric' })
}
