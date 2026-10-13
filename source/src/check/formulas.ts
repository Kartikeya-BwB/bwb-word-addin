// Chemical formulas written with digits (CO2, CH4, N2O) and whether their digits are set as subscript, read from the
// body's WordprocessingML. Office.js has no per-character formatting, so this uses the deep scan's parsed package.
// Paragraphs are numbered as runFormatting.ts numbers them: in document order, leaving out text boxes and fallbacks.

import type { FormulaFormatting } from './types.ts'

const W = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main'
const PKG = 'http://schemas.microsoft.com/office/2006/xmlPackage'
const MC = 'http://schemas.openxmlformats.org/markup-compatibility/2006'

/** CO2, CH4, N2O, SF6, NF3, SO2 and NO2, with an optional tonnes prefix (tCO2, MtCO2) and the e of CO2e. */
export const DIGIT_FORMULA = /(?<![\p{L}\p{N}])(?:[kMG]?t)?(?:CO2|CH4|N2O|SF6|NF3|SO2|NO2)e?(?![\p{L}\p{N}])/gu

function isW(node: Node | null, name: string): node is Element {
  return node !== null && node.nodeType === 1 && (node as Element).namespaceURI === W && (node as Element).localName === name
}

function hidden(el: Element): boolean {
  for (let a: Node | null = el.parentNode; a; a = a.parentNode) {
    if (a.nodeType !== 1) continue
    const e = a as Element
    if ((e.namespaceURI === W && e.localName === 'txbxContent') || (e.namespaceURI === MC && e.localName === 'Fallback')) return true
  }
  return false
}

function bodyRoot(pkg: Document): Element | Document {
  const parts = pkg.getElementsByTagNameNS(PKG, 'part')
  for (let i = 0; i < parts.length; i += 1) {
    const part = parts.item(i)
    if (part?.getAttributeNS(PKG, 'name') === '/word/document.xml' || part?.getAttribute('pkg:name') === '/word/document.xml') return part
  }
  return pkg
}

/** The paragraph's text with a subscript flag for each character, from its own runs (not those of nested text boxes). */
function textWithSubscripts(p: Element): { text: string; sub: boolean[] } {
  let text = ''
  const sub: boolean[] = []
  const runs = p.getElementsByTagNameNS(W, 'r')
  for (let j = 0; j < runs.length; j += 1) {
    const r = runs.item(j)
    if (!r) continue
    let owner: Node | null = r.parentNode
    while (owner && !isW(owner, 'p')) owner = owner.parentNode
    if (owner !== p) continue
    let subscript = false
    for (let c = r.firstChild; c; c = c.nextSibling) {
      if (!isW(c, 'rPr')) continue
      for (let v = c.firstChild; v; v = v.nextSibling) {
        if (isW(v, 'vertAlign')) subscript = (v.getAttributeNS(W, 'val') || v.getAttribute('w:val')) === 'subscript'
      }
    }
    for (let c = r.firstChild; c; c = c.nextSibling) {
      if (!isW(c, 't')) continue
      const t = c.textContent ?? ''
      text += t
      for (let k = 0; k < t.length; k += 1) sub.push(subscript)
    }
  }
  return { text, sub }
}

/** For each paragraph that writes a formula with digits, how many have the digits in subscript and how many do not. */
export function scanFormulaFormatting(pkg: Document, trusted = true): FormulaFormatting[] {
  const all = bodyRoot(pkg).getElementsByTagNameNS(W, 'p')
  const out: FormulaFormatting[] = []
  let index = 0
  for (let i = 0; i < all.length; i += 1) {
    const p = all.item(i)
    if (!p || hidden(p)) continue
    const quick = p.textContent ?? ''
    if (/(?:CO2|CH4|N2O|SF6|NF3|SO2|NO2)/.test(quick)) {
      const { text, sub } = textWithSubscripts(p)
      let subscript = 0
      let plain = 0
      for (const m of text.matchAll(DIGIT_FORMULA)) {
        const digits = [...m[0]].map((ch, k) => (/\d/.test(ch) ? m.index + k : -1)).filter((k) => k >= 0)
        if (digits.every((k) => sub[k])) subscript += 1
        else plain += 1
      }
      if (subscript || plain) out.push({ paragraph: trusted ? index : null, subscript, plain })
    }
    index += 1
  }
  return out
}
