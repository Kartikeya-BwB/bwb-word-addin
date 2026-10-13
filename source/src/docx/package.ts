/** Minimal OPC package handling for .docx files on fflate (read, edit parts and relationships, prune, write): a port of
 * bwb_lib/package.py. Relationship files are rewritten in the Python format, so the output matches part for part. */
import { strFromU8, strToU8, unzipSync, zipSync, type Zippable } from 'fflate'
import { esc, XML_DECL } from './xml.ts'

export const REL_NS = 'http://schemas.openxmlformats.org/package/2006/relationships'
export const RT = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/'
const CORE_REL = 'http://schemas.openxmlformats.org/package/2006/relationships/metadata/core-properties'

export const CT_DEFAULTS: Readonly<Record<string, string>> = {
  rels: 'application/vnd.openxmlformats-package.relationships+xml',
  xml: 'application/xml',
  png: 'image/png',
  jpg: 'image/jpeg',
  jpeg: 'image/jpeg',
  gif: 'image/gif',
  svg: 'image/svg+xml',
  emf: 'image/x-emf',
  wmf: 'image/x-wmf',
  bmp: 'image/bmp',
  tif: 'image/tiff',
  tiff: 'image/tiff',
}
const WML = 'application/vnd.openxmlformats-officedocument.wordprocessingml.'
export const CT_OVERRIDES: ReadonlyArray<readonly [string, string]> = [
  ['word/document.xml', WML + 'document.main+xml'],
  ['word/styles.xml', WML + 'styles+xml'],
  ['word/numbering.xml', WML + 'numbering+xml'],
  ['word/settings.xml', WML + 'settings+xml'],
  ['word/webSettings.xml', WML + 'webSettings+xml'],
  ['word/footnotes.xml', WML + 'footnotes+xml'],
  ['word/endnotes.xml', WML + 'endnotes+xml'],
  ['word/fontTable.xml', WML + 'fontTable+xml'],
  ['word/theme/theme1.xml', 'application/vnd.openxmlformats-officedocument.theme+xml'],
  ['docProps/core.xml', 'application/vnd.openxmlformats-package.core-properties+xml'],
  ['docProps/app.xml', 'application/vnd.openxmlformats-officedocument.extended-properties+xml'],
]
const STALE = ['docProps/custom.xml', 'customXml/item1.xml', 'customXml/itemProps1.xml', 'customXml/_rels/item1.xml.rels']
const CUSTOM_CT = 'application/vnd.openxmlformats-officedocument.custom-properties+xml'
const CUSTOM_FMTID = '{D5CDD505-2E9C-101B-9397-08002B2CF9AE}'

export interface Rel {
  id: string
  type: string
  target: string
  /** 'External' for hyperlinks, null for parts in the package. */
  mode: string | null
}

const REL_RE = /<(?:[\w.-]+:)?Relationship\b([^>]*?)\/?>/g
const ATTR_RE = /([\w:.-]+)\s*=\s*(?:"([^"]*)"|'([^']*)')/g
const ENTITIES: Readonly<Record<string, string>> = { lt: '<', gt: '>', quot: '"', apos: "'", amp: '&' }

export function unescapeXml(s: string): string {
  return s.replace(/&(#x[0-9a-fA-F]+|#\d+|lt|gt|quot|apos|amp);/g, (_m, e: string) => {
    if (e.startsWith('#x')) return String.fromCodePoint(parseInt(e.slice(2), 16))
    if (e.startsWith('#')) return String.fromCodePoint(parseInt(e.slice(1), 10))
    return ENTITIES[e] ?? ''
  })
}

export function parseRels(xml: string): Rel[] {
  const out: Rel[] = []
  for (const m of xml.matchAll(REL_RE)) {
    const a: Record<string, string> = {}
    for (const am of (m[1] ?? '').matchAll(ATTR_RE)) a[am[1] ?? ''] = unescapeXml(am[2] ?? am[3] ?? '')
    out.push({ id: a['Id'] ?? '', type: a['Type'] ?? '', target: a['Target'] ?? '', mode: a['TargetMode'] ?? null })
  }
  return out
}

export function relsXml(rels: readonly Rel[]): string {
  let out = XML_DECL + `<Relationships xmlns="${REL_NS}">`
  for (const r of rels) {
    const mode = r.mode ? ` TargetMode="${r.mode}"` : ''
    out += `<Relationship Id="${r.id}" Type="${r.type}" Target="${esc(r.target)}"${mode}/>`
  }
  return out + '</Relationships>'
}

function byCodePoint(a: string, b: string): number {
  return a < b ? -1 : a > b ? 1 : 0
}

export interface CoreProps {
  title: string
  subject: string
  keywords: string
  description: string
  creator?: string
}

function isoSeconds(d: Date): string {
  return d.toISOString().replace(/\.\d{3}Z$/, 'Z')
}

export function coreXml(p: CoreProps, now: Date): string {
  const creator = p.creator ?? 'Bankers without Boundaries'
  const t = isoSeconds(now)
  return (
    XML_DECL +
    '<cp:coreProperties xmlns:cp="http://schemas.openxmlformats.org/package/2006/metadata/core-properties" ' +
    'xmlns:dc="http://purl.org/dc/elements/1.1/" xmlns:dcterms="http://purl.org/dc/terms/" ' +
    'xmlns:dcmitype="http://purl.org/dc/dcmitype/" xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance">' +
    `<dc:title>${esc(p.title)}</dc:title><dc:subject>${esc(p.subject)}</dc:subject><dc:creator>${esc(creator)}</dc:creator>` +
    `<cp:keywords>${esc(p.keywords)}</cp:keywords><dc:description>${esc(p.description)}</dc:description>` +
    `<cp:lastModifiedBy>${esc(creator)}</cp:lastModifiedBy><cp:revision>1</cp:revision>` +
    `<dcterms:created xsi:type="dcterms:W3CDTF">${t}</dcterms:created>` +
    `<dcterms:modified xsi:type="dcterms:W3CDTF">${t}</dcterms:modified></cp:coreProperties>`
  )
}

export function appXml(): string {
  return (
    XML_DECL +
    '<Properties xmlns="http://schemas.openxmlformats.org/officeDocument/2006/extended-properties" ' +
    'xmlns:vt="http://schemas.openxmlformats.org/officeDocument/2006/docPropsVTypes">' +
    '<Application>Microsoft Office Word</Application><Company>Bankers without Boundaries</Company>' +
    '<AppVersion>16.0000</AppVersion></Properties>'
  )
}

/** Office custom document properties: text values, property ids from 2. */
export function customXml(props: Readonly<Record<string, string>>): string {
  let pid = 2
  let items = ''
  for (const [name, value] of Object.entries(props)) {
    items += `<property fmtid="${CUSTOM_FMTID}" pid="${pid++}" name="${esc(name)}"><vt:lpwstr>${esc(value)}</vt:lpwstr></property>`
  }
  return (
    XML_DECL +
    '<Properties xmlns="http://schemas.openxmlformats.org/officeDocument/2006/custom-properties" ' +
    `xmlns:vt="http://schemas.openxmlformats.org/officeDocument/2006/docPropsVTypes">${items}</Properties>`
  )
}

export interface SaveOptions {
  core?: CoreProps | null
  prune?: boolean
  now?: Date
  customProperties?: Readonly<Record<string, string>> | undefined
}

const STORED = new Set(['png', 'jpg', 'jpeg', 'gif'])

export class Package {
  private readonly files = new Map<string, Uint8Array | string>()
  private ridCounter = 1000

  constructor(files?: Iterable<readonly [string, Uint8Array | string]>) {
    if (files) for (const [k, v] of files) this.files.set(k, v)
  }

  static open(bytes: Uint8Array): Package {
    const pkg = new Package()
    for (const [name, data] of Object.entries(unzipSync(bytes))) if (!name.endsWith('/')) pkg.files.set(name, data)
    return pkg
  }

  names(): string[] {
    return [...this.files.keys()]
  }

  has(name: string): boolean {
    return this.files.has(name)
  }

  get(name: string): Uint8Array {
    const v = this.files.get(name)
    if (v === undefined) throw new Error(`The package has no part ${name}.`)
    return typeof v === 'string' ? strToU8(v) : v
  }

  text(name: string): string {
    const v = this.files.get(name)
    if (v === undefined) throw new Error(`The package has no part ${name}.`)
    return typeof v === 'string' ? v : strFromU8(v)
  }

  set(name: string, data: Uint8Array | string): void {
    this.files.set(name, data)
  }

  remove(name: string): void {
    this.files.delete(name)
  }

  // ---- relationships ----
  relsPath(part: string): string {
    const i = part.lastIndexOf('/')
    return `${part.slice(0, i)}/_rels/${part.slice(i + 1)}.rels`
  }

  rels(part = 'word/document.xml'): { path: string; rels: Rel[] } {
    const path = this.relsPath(part)
    return { path, rels: this.files.has(path) ? parseRels(this.text(path)) : [] }
  }

  writeRels(path: string, rels: readonly Rel[]): void {
    this.files.set(path, relsXml(rels))
  }

  relIdFor(target: string, part = 'word/document.xml'): string | null {
    for (const r of this.rels(part).rels) if (r.target === target) return r.id
    return null
  }

  addRel(rtype: string, target: string, external = false, part = 'word/document.xml'): string {
    const { path, rels } = this.rels(part)
    for (const r of rels) if (r.target === target && r.type.endsWith(rtype)) return r.id
    const used = new Set(rels.map((r) => r.id))
    while (used.has(`rId${this.ridCounter}`)) this.ridCounter += 1
    const rid = `rId${this.ridCounter}`
    this.ridCounter += 1
    rels.push({ id: rid, type: RT + rtype, target, mode: external ? 'External' : null })
    this.writeRels(path, rels)
    return rid
  }

  /** Store an image under word/media/<name> and return its relationship id (the existing one if present). */
  addMedia(name: string, data: Uint8Array | string): string {
    this.files.set(`word/media/${name}`, data)
    return this.addRel('image', `media/${name}`)
  }

  /** Replace all relationships of a part. items: [Id, type suffix or full type, target]. */
  setRels(part: string, items: ReadonlyArray<readonly [string, string, string]>): void {
    this.writeRels(
      this.relsPath(part),
      items.map(([id, t, target]) => ({ id, type: t.startsWith('http') ? t : RT + t, target, mode: null })),
    )
  }

  // ---- pruning ----
  /** Drop image and hyperlink relationships the document no longer references, then media nothing points to. */
  prune(): void {
    const doc = this.text('word/document.xml')
    const used = new Set<string>()
    for (const m of doc.matchAll(/r:(?:embed|id|link|pict)="([^"]+)"/g)) used.add(m[1] ?? '')
    const { path, rels } = this.rels()
    this.writeRels(
      path,
      rels.filter((r) => {
        const t = r.type.slice(r.type.lastIndexOf('/') + 1)
        return !((t === 'image' || t === 'hyperlink') && !used.has(r.id))
      }),
    )
    const referenced = new Set<string>()
    for (const p of [...this.files.keys()]) {
      if (!p.endsWith('.rels')) continue
      const flat = p.replaceAll('_rels/', '')
      const base = flat.slice(0, flat.lastIndexOf('.rels'))
      const bdir = base.includes('/') ? base.slice(0, base.lastIndexOf('/')) : ''
      for (const r of parseRels(this.text(p))) {
        if (r.mode === 'External' || !r.target) continue
        const t = r.target
        const full = t.startsWith('/') ? t.replace(/^\/+/, '') : bdir ? `${bdir}/${t}` : t
        const parts: string[] = []
        for (const seg of full.split('/')) {
          if (seg === '..') parts.pop()
          else parts.push(seg)
        }
        referenced.add(parts.join('/'))
      }
    }
    for (const p of [...this.files.keys()]) if (p.startsWith('word/media/') && !referenced.has(p)) this.files.delete(p)
  }

  // ---- content types and package relationships ----
  contentTypes(): string {
    const exts = new Set<string>()
    for (const p of this.files.keys()) {
      if (p.slice(p.lastIndexOf('/') + 1).includes('.')) exts.add(p.slice(p.lastIndexOf('.') + 1).toLowerCase())
    }
    let out = XML_DECL + '<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">'
    for (const e of [...exts].sort(byCodePoint)) {
      const ct = Object.hasOwn(CT_DEFAULTS, e) ? CT_DEFAULTS[e] : undefined
      if (ct) out += `<Default Extension="${e}" ContentType="${ct}"/>`
    }
    for (const [part, ct] of CT_OVERRIDES) if (this.files.has(part)) out += `<Override PartName="/${part}" ContentType="${ct}"/>`
    for (const p of [...this.files.keys()].sort(byCodePoint)) {
      const m = /^word\/(header|footer)\d+\.xml$/.exec(p)
      if (m) out += `<Override PartName="/${p}" ContentType="${WML}${m[1] ?? ''}+xml"/>`
    }
    return out + '</Types>'
  }

  /** The .docx bytes, written as Package.save does: prune, content types, package relationships, properties. */
  save(o: SaveOptions = {}): Uint8Array {
    const now = o.now ?? new Date()
    if (o.prune ?? true) this.prune()
    this.files.set('[Content_Types].xml', this.contentTypes())
    this.files.set(
      '_rels/.rels',
      XML_DECL +
        `<Relationships xmlns="${REL_NS}">` +
        `<Relationship Id="rId1" Type="${RT}officeDocument" Target="word/document.xml"/>` +
        `<Relationship Id="rId2" Type="${CORE_REL}" Target="docProps/core.xml"/>` +
        `<Relationship Id="rId3" Type="${RT}extended-properties" Target="docProps/app.xml"/>` +
        '</Relationships>',
    )
    for (const stale of STALE) this.files.delete(stale)
    if (o.core) {
      this.files.set('docProps/core.xml', coreXml(o.core, now))
      this.files.set('docProps/app.xml', appXml())
    }
    const custom = o.customProperties
    if (custom && Object.keys(custom).length) {
      // an add-in extension; without custom properties the package is exactly the Python one
      this.files.set('docProps/custom.xml', customXml(custom))
      this.files.set(
        '[Content_Types].xml',
        this.text('[Content_Types].xml').replace(
          '</Types>',
          `<Override PartName="/docProps/custom.xml" ContentType="${CUSTOM_CT}"/></Types>`,
        ),
      )
      this.files.set(
        '_rels/.rels',
        this.text('_rels/.rels').replace(
          '</Relationships>',
          `<Relationship Id="rId4" Type="${RT}custom-properties" Target="docProps/custom.xml"/></Relationships>`,
        ),
      )
    }
    const order = ['[Content_Types].xml', '_rels/.rels']
    const names = [...order, ...[...this.files.keys()].filter((n) => !order.includes(n)).sort(byCodePoint)]
    const zip: Zippable = {}
    for (const n of names) {
      const ext = n.slice(n.lastIndexOf('.') + 1).toLowerCase()
      zip[n] = [this.get(n), { level: STORED.has(ext) ? 0 : 6, mtime: now }]
    }
    return zipSync(zip)
  }
}
