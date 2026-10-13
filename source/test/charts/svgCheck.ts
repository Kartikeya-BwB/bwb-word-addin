// Reads chart SVG back the way a reviewer would: is it well-formed XML, which colours and text sizes does it use,
// and where does each piece of text sit.

import { DOMParser } from '@xmldom/xmldom'
import type { Element as XmlElement } from '@xmldom/xmldom'
import { textWidth } from '../../src/charts/measure.ts'

export interface ParsedSvg {
  root: XmlElement
  errors: string[]
  width: number
  height: number
}

export function parseSvg(svg: string): ParsedSvg {
  const errors: string[] = []
  const doc = new DOMParser({ onError: (level, message) => errors.push(`${level}: ${message}`) }).parseFromString(svg, 'image/svg+xml')
  const root = doc.documentElement
  if (!root) throw new Error('No root element')
  const [, , width = 0, height = 0] = (root.getAttribute('viewBox') ?? '').split(/\s+/).map(Number)
  return { root, errors, width, height }
}

export function elements(root: XmlElement): XmlElement[] {
  const out: XmlElement[] = []
  const walk = (node: XmlElement): void => {
    out.push(node)
    for (let i = 0; i < node.childNodes.length; i += 1) {
      const child = node.childNodes[i]
      if (child && child.nodeType === 1) walk(child as XmlElement)
    }
  }
  walk(root)
  return out
}

/** Every colour used in fill and stroke attributes, upper-case hex without '#'. */
export function colours(root: XmlElement): Set<string> {
  const found = new Set<string>()
  for (const el of elements(root)) {
    for (const name of ['fill', 'stroke', 'stop-color', 'color']) {
      const value = el.getAttribute(name)
      if (value && value !== 'none') found.add(value.replace('#', '').toUpperCase())
    }
    if (el.getAttribute('style')) found.add(`style:${el.getAttribute('style') ?? ''}`)
  }
  return found
}

export interface TextBox {
  text: string
  size: number
  x0: number
  x1: number
  y0: number
  y1: number
}

/** Text elements with the box their glyphs take, from the same width estimate the layout uses. */
export function textBoxes(root: XmlElement): TextBox[] {
  return elements(root)
    .filter((el) => el.tagName === 'text')
    .map((el) => {
      const content = el.textContent ?? ''
      const size = Number(el.getAttribute('font-size'))
      const bold = el.getAttribute('font-weight') === 'bold'
      const width = textWidth(content, size, bold)
      const x = Number(el.getAttribute('x'))
      const y = Number(el.getAttribute('y'))
      const anchor = el.getAttribute('text-anchor') ?? 'start'
      const x0 = anchor === 'middle' ? x - width / 2 : anchor === 'end' ? x - width : x
      return { text: content, size, x0, x1: x0 + width, y0: y - size * 0.72, y1: y + size * 0.22 }
    })
}

/** Pairs of text boxes that overlap by more than a hair. */
export function overlaps(boxes: TextBox[], tolerance = 0.25): Array<[string, string]> {
  const out: Array<[string, string]> = []
  boxes.forEach((a, i) => {
    boxes.slice(i + 1).forEach((b) => {
      const w = Math.min(a.x1, b.x1) - Math.max(a.x0, b.x0)
      const h = Math.min(a.y1, b.y1) - Math.max(a.y0, b.y0)
      if (w > tolerance && h > tolerance) out.push([a.text, b.text])
    })
  })
  return out
}
