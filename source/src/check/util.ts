import type { ParagraphInfo } from './types.ts'

const BUILT_IN_HEADING = /^Heading([1-9])$/

/** Heading level of a paragraph from its style, or null for body text. */
export function headingLevel(p: ParagraphInfo): number | null {
  const builtIn = BUILT_IN_HEADING.exec(p.styleBuiltIn)
  if (builtIn) return Number(builtIn[1])
  const named = /^(?:heading ([1-9])|BwB Heading ([1-3]) Unnumbered)$/i.exec(p.style.trim())
  if (named) return Number(named[1] ?? named[2])
  return null
}

export function isTitle(p: ParagraphInfo): boolean {
  return p.styleBuiltIn === 'Title' || p.style === 'BwB Document Title'
}

export function isTocEntry(p: ParagraphInfo): boolean {
  return /^toc\s?\d$/i.test(p.style) || /^Toc\d$/.test(p.styleBuiltIn)
}

/** Paragraphs whose words are quoted or generated, so the writing rules leave them alone. */
export function isVerbatim(p: ParagraphInfo): boolean {
  return isTocEntry(p) || p.style === 'BwB Pull Quote' || p.style === 'BwB Pull Quote Source'
}

export function wordCount(text: string): number {
  const words = text.trim().match(/\S+/g)
  return words ? words.length : 0
}

export function excerpt(text: string, start = 0, end = text.length, pad = 30): string {
  const a = Math.max(0, start - pad)
  const b = Math.min(text.length, end + pad)
  return `${a > 0 ? '…' : ''}${text.slice(a, b).replace(/\s+/g, ' ').trim()}${b < text.length ? '…' : ''}`
}

export function plural(n: number, one: string, many = `${one}s`): string {
  return `${n} ${n === 1 ? one : many}`
}
