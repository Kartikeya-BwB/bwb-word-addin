import { existsSync, readFileSync } from 'node:fs'
import { join } from 'node:path'
import { DOMParser } from '@xmldom/xmldom'
import { describe, expect, it } from 'vitest'

const ROOT = join(import.meta.dirname, '..')
const BT = 'http://schemas.microsoft.com/office/officeappbasictypes/1.0'

function load(file: string) {
  const xml = readFileSync(join(ROOT, file), 'utf8')
  const doc = new DOMParser().parseFromString(xml, 'application/xml')
  const root = doc.documentElement
  if (!root) throw new Error(`${file} has no root element`)
  return { xml, root }
}

describe.each(['public/manifest.xml', 'manifest.dev.xml'])('%s', (file) => {
  const { xml, root } = load(file)

  it('uses the schema namespaces, spelled exactly', () => {
    expect(root.namespaceURI).toBe('http://schemas.microsoft.com/office/appforoffice/1.1')
    expect(root.getAttribute('xmlns:bt')).toBe(BT)
    expect(root.getElementsByTagNameNS(BT, 'Image').length).toBeGreaterThan(0)
  })

  it('defines every resource it refers to, and refers to every resource it defines', () => {
    const referenced = new Set([...xml.matchAll(/resid="([^"]+)"/g)].map((m) => m[1]))
    const defined = [...xml.matchAll(/<bt:(?:Image|Url|String) id="([^"]+)"/g)].map((m) => m[1])
    expect(new Set(defined).size).toBe(defined.length)
    expect([...referenced].filter((r) => !defined.includes(r))).toEqual([])
    expect(defined.filter((d) => !referenced.has(d))).toEqual([])
    for (const id of defined) expect(id?.length ?? 0).toBeLessThanOrEqual(32)
  })

  it('points only at icons that exist in public/', () => {
    const base = /<SourceLocation DefaultValue="([^"]+)index\.html"/.exec(xml)?.[1] ?? ''
    const images = [...xml.matchAll(/DefaultValue="([^"]+\.png)"/g)].map((m) => m[1] ?? '')
    expect(images.length).toBeGreaterThan(100)
    for (const url of images) expect(existsSync(join(ROOT, 'public', url.slice(base.length))), url).toBe(true)
  })
})

it('keeps the published add-in ID stable', () => {
  expect(load('public/manifest.xml').xml).toContain('<Id>b0732c64-7b19-470c-8f8f-22e8d6799c0f</Id>')
})
