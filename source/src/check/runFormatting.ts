// Reads direct run formatting (fonts, colours, highlighting, sizes) and inline pictures from the body's WordprocessingML.
// Office.js only reports formatting per paragraph, and a mixed paragraph reports nothing, so this finds the
// off-brand runs the API cannot see. Paragraphs are numbered in document order, as body.paragraphs is.

import type { RunFormatting } from './types.ts'

const W = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main'
const PKG = 'http://schemas.microsoft.com/office/2006/xmlPackage'
const MC = 'http://schemas.openxmlformats.org/markup-compatibility/2006'

export type XmlParser = (xml: string) => Document

function documentPart(pkg: Document): Element | null {
  const parts = pkg.getElementsByTagNameNS(PKG, 'part')
  for (let i = 0; i < parts.length; i += 1) {
    const part = parts.item(i)
    if (part?.getAttributeNS(PKG, 'name') === '/word/document.xml' || part?.getAttribute('pkg:name') === '/word/document.xml') {
      return part
    }
  }
  return null
}

/** True for paragraphs Office.js does not list in body.paragraphs: text boxes and fallback copies. */
function hidden(el: Element): boolean {
  for (let a: Node | null = el.parentNode; a; a = a.parentNode) {
    if (a.nodeType !== 1) continue
    const e = a as Element
    if (e.namespaceURI === W && e.localName === 'txbxContent') return true
    if (e.namespaceURI === MC && e.localName === 'Fallback') return true
  }
  return false
}

function attr(el: Element | null | undefined, name: string): string | null {
  if (!el) return null
  return el.getAttributeNS(W, name) || el.getAttribute(`w:${name}`)
}

function firstChild(el: Element, name: string): Element | null {
  for (let c = el.firstChild; c; c = c.nextSibling) {
    if (c.nodeType === 1 && (c as Element).namespaceURI === W && (c as Element).localName === name) return c as Element
  }
  return null
}

export interface RunFormattingResult {
  paragraphs: RunFormatting[]
  /** Paragraphs counted in the XML; equal to the API's paragraph count when locations can be trusted. */
  counted: number
}

export function scanRunFormatting(flatOpc: string, parse: XmlParser): RunFormattingResult | null {
  const pkg = parse(flatOpc)
  const part = documentPart(pkg)
  const root = part ?? pkg
  const all = root.getElementsByTagNameNS(W, 'p')
  const out: RunFormatting[] = []
  let index = 0
  for (let i = 0; i < all.length; i += 1) {
    const p = all.item(i)
    if (!p || hidden(p)) continue
    const fonts = new Set<string>()
    const colours = new Set<string>()
    const highlights = new Set<string>()
    let smallest: number | null = null
    let pictures = 0
    const runs = p.getElementsByTagNameNS(W, 'r')
    for (let j = 0; j < runs.length; j += 1) {
      const r = runs.item(j)
      if (!r || r.parentNode === null) continue
      // Runs of nested paragraphs (inside text boxes) belong to those paragraphs, not this one.
      let owner: Node | null = r.parentNode
      while (owner && !((owner as Element).namespaceURI === W && (owner as Element).localName === 'p')) owner = owner.parentNode
      if (owner !== p) continue
      if (firstChild(r, 'drawing') || firstChild(r, 'pict')) pictures += 1
      const rPr = firstChild(r, 'rPr')
      if (!rPr) continue
      const text = r.getElementsByTagNameNS(W, 't')
      if (text.length === 0) continue
      const font = attr(firstChild(rPr, 'rFonts'), 'ascii')
      if (font) fonts.add(font)
      const colour = attr(firstChild(rPr, 'color'), 'val')
      if (colour && colour !== 'auto') colours.add(colour.toUpperCase())
      const highlight = attr(firstChild(rPr, 'highlight'), 'val')
      if (highlight && highlight !== 'none') highlights.add(highlight)
      const shading = attr(firstChild(rPr, 'shd'), 'fill')
      if (shading && shading !== 'auto' && /^[0-9A-Fa-f]{6}$/.test(shading)) highlights.add(`#${shading.toUpperCase()}`)
      const size = Number(attr(firstChild(rPr, 'sz'), 'val'))
      if (Number.isFinite(size) && size > 0) smallest = smallest === null ? size : Math.min(smallest, size)
    }
    if (fonts.size || colours.size || highlights.size || smallest !== null || pictures) {
      out.push({ paragraph: index, fonts: [...fonts], colours: [...colours], highlights: [...highlights], smallestHalfPoints: smallest, pictures })
    }
    index += 1
  }
  return { paragraphs: out, counted: index }
}
