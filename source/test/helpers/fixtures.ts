/**
 * Node-only access to the differential fixtures: the content files in test/fixtures/docx, the Python builder's outputs
 * in test/fixtures/docx/expected (written by scripts/gen-docx-fixtures.py) and the committed template assets.
 */
import { readdirSync, readFileSync } from 'node:fs'
import { join } from 'node:path'
import { fileURLToPath } from 'node:url'
import { readImageInfo } from '../../src/docx/imageInfo.ts'
import type { Block, Content, PictureData } from '../../src/docx/types.ts'

export const SOURCE_DIR = fileURLToPath(new URL('../../', import.meta.url))
export const FIXTURES_DIR = join(SOURCE_DIR, 'test', 'fixtures', 'docx')
export const EXPECTED_DIR = join(FIXTURES_DIR, 'expected')

/** The date gen-docx-fixtures.py gives the Python builder (5 October 2026). */
export const FIXED_NOW = new Date(2026, 9, 5, 12, 0, 0)

/** The XML parts the fixture generator saves in full, so a failure shows the difference. */
export const SAVED_PARTS = [
  'word/document.xml',
  'word/numbering.xml',
  'word/footnotes.xml',
  'word/_rels/document.xml.rels',
  '[Content_Types].xml',
  'docProps/core.xml',
] as const

export interface Manifest {
  source: string
  parts: Record<string, { sha256: string; size: number }>
  media: string[]
  icons: string[]
  warnings: string[]
}

export function readBytes(...path: string[]): Uint8Array {
  return new Uint8Array(readFileSync(join(SOURCE_DIR, ...path)))
}

export function baseTemplate(): Uint8Array {
  return readBytes('public', 'templates', 'bwb-base.docx')
}

export function coverPhoto(): Uint8Array {
  return readBytes('public', 'templates', 'bwb-cover.jpg')
}

export function fixtureNames(): string[] {
  return readdirSync(FIXTURES_DIR)
    .filter((f) => f.endsWith('.json'))
    .map((f) => f.slice(0, -'.json'.length))
    .sort()
}

/** A picture from the fixtures folder, as the Python builder's load_image passes it on (bytes unchanged). */
export function picture(relPath: string): PictureData {
  const data = new Uint8Array(readFileSync(join(FIXTURES_DIR, relPath)))
  const info = readImageInfo(data)
  if (!info) throw new Error(`${relPath} is not a PNG, JPEG or GIF image`)
  return { data, ...info }
}

type JsonObject = Record<string, unknown>

/** The fixture's content with figure and cover image paths replaced by picture bytes. */
export function loadFixture(name: string): Content {
  const raw = JSON.parse(readFileSync(join(FIXTURES_DIR, `${name}.json`), 'utf8')) as JsonObject
  const meta = (raw['meta'] ?? {}) as JsonObject
  if (typeof meta['cover_image'] === 'string') meta['cover_image'] = picture(meta['cover_image'])
  const blocks = ((raw['blocks'] ?? []) as JsonObject[]).map((b) => {
    if (b['type'] !== 'figure' || typeof b['path'] !== 'string') return b
    const { path, ...rest } = b
    return { ...rest, ...picture(path as string) }
  })
  return { meta, blocks: blocks as unknown as Block[] } as Content
}

export function loadManifest(name: string): Manifest {
  return JSON.parse(readFileSync(join(EXPECTED_DIR, name, 'manifest.json'), 'utf8')) as Manifest
}

export function expectedPart(name: string, part: string): string {
  return readFileSync(join(EXPECTED_DIR, name, part), 'utf8')
}

/** Where two strings first differ, with some context, for readable failures on long single-line XML. */
export function firstDifference(actual: string, expected: string, context = 160): string | null {
  if (actual === expected) return null
  let i = 0
  while (i < actual.length && i < expected.length && actual[i] === expected[i]) i++
  const from = Math.max(0, i - context)
  return (
    `first difference at character ${i} (lengths ${actual.length} and ${expected.length})\n` +
    `  actual:   ...${actual.slice(from, i + context)}\n` +
    `  expected: ...${expected.slice(from, i + context)}`
  )
}
