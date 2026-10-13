import { readFileSync } from 'node:fs'
import { join } from 'node:path'
import { strFromU8, unzipSync } from 'fflate'
import { afterEach, beforeEach, describe, expect, it, vi } from 'vitest'
import { DOC_TYPES, findDocType } from '../../src/brand/docTypes.ts'
import type { NewDocumentRequest } from '../../src/word/newDocument.ts'
import { bwbDate, createNewDocument } from '../../src/word/newDocument.ts'
import { createModel } from '../fakeWord/model.ts'
import { installFakeWord, uninstallFakeWord } from '../fakeWord/install.ts'

const PUBLIC = join(import.meta.dirname, '../../public')

beforeEach(() => {
  // The pane fetches the template and cover photo from its own site; serve them from public/.
  vi.stubGlobal('fetch', async (url: string) => {
    if (url.startsWith('./templates/')) return new Response(readFileSync(join(PUBLIC, url.slice(2))))
    return new Response('', { status: 404 })
  })
})
afterEach(() => {
  vi.unstubAllGlobals()
  uninstallFakeWord()
})

function request(typeId: string, over: Partial<NewDocumentRequest> = {}): NewDocumentRequest {
  const type = findDocType(typeId)
  if (!type) throw new Error(`no type ${typeId}`)
  return {
    type, title: 'Blended finance for adaptation', tagline: 'Lessons from five programmes', date: '13 October 2026',
    preparedBy: '', details: [], disclaimer: true, project: 'NZC', deliverable: 'D3.2', funder: '', ...over,
  }
}

async function created(model: ReturnType<typeof createModel>): Promise<Record<string, string>> {
  const b64 = model.created[0]
  if (!b64) throw new Error('no document was opened')
  const files = unzipSync(new Uint8Array(Buffer.from(b64, 'base64')))
  return Object.fromEntries(Object.entries(files).filter(([n]) => n.endsWith('.xml') || n.endsWith('.rels')).map(([n, d]) => [n, strFromU8(d)]))
}

describe('createNewDocument', () => {
  it('opens a final report with its cover, structure, disclaimer and BwB properties', async () => {
    const model = createModel()
    installFakeWord(model)
    await createNewDocument(request('final-report', { funder: 'eu-grant' }))
    const files = await created(model)
    const doc = files['word/document.xml'] ?? ''
    expect(doc).toContain('Blended finance for adaptation')
    expect(doc).toContain('Lessons from five programmes')
    expect(doc).toContain('Funded by the European Union')
    expect(doc).toContain('953637')
    expect(doc).toMatch(/w:instrText[^>]*>[^<]*TOC/)
    const custom = files['docProps/custom.xml'] ?? ''
    expect(custom).toContain('BwB document type')
    expect(custom).toContain('final-report')
    expect(custom).toContain('NZC')
    expect(files['docProps/core.xml']).toContain('Blended finance for adaptation')
  })

  it('opens a memo with its details table and no disclaimer when it is turned off', async () => {
    const model = createModel()
    installFakeWord(model)
    await createNewDocument(request('memo', { disclaimer: false, details: [['To', 'The board'], ['From', 'Programme team']] }))
    const doc = (await created(model))['word/document.xml'] ?? ''
    expect(doc).toContain('The board')
    expect(doc).toContain('Programme team')
    expect(doc).not.toContain('953637')
  })

  it('builds every document type without errors', async () => {
    for (const type of DOC_TYPES) {
      const model = createModel()
      installFakeWord(model)
      const { warnings } = await createNewDocument(request(type.id))
      expect(model.created, type.id).toHaveLength(1)
      expect(warnings.filter((w) => !/icon/i.test(w)), type.id).toEqual([])
      uninstallFakeWord()
    }
  })
})

describe('bwbDate', () => {
  it('writes dates as D Month YYYY', () => {
    expect(bwbDate(new Date(2026, 9, 5))).toBe('5 October 2026')
  })
})
