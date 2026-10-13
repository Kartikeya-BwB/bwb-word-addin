// Brand: the BwB template, fonts, palette, lime on light backgrounds, highlighting and very small text.

import { ALLOWED_FONTS, ALLOWED_HEX, COLOURS, MIN_TEXT_PT } from '../brand/tokens.ts'
import type { DocSnapshot, Finding, Location, ParagraphInfo } from './types.ts'
import { excerpt, headingLevel } from './util.ts'

const hex = (colour: string | null): string | null => {
  if (!colour) return null
  const m = /^#?([0-9a-f]{6})$/i.exec(colour.trim())
  return m?.[1] ? m[1].toUpperCase() : null
}

const where = (p: ParagraphInfo): Location => ({ kind: 'paragraph', paragraph: p.index, expectedText: p.text })

function templateFindings(snapshot: DocSnapshot): Finding[] {
  if (snapshot.bwbStyles !== false || snapshot.info.clientTemplate) return []
  return [{
    id: 'brand:template', rule: 'brand/not-template', group: 'brand', severity: 'minor',
    title: 'Not built on the BwB template',
    detail: 'Add the BwB styles, or start again from New document. If the client’s template applies, record that on the Start tab.',
    fix: { kind: 'add-styles' }, fixLabel: 'Add BwB styles and restyle Word’s own',
  }]
}

function fontFindings(snapshot: DocSnapshot): Finding[] {
  const text = snapshot.paragraphs.filter((p) => p.text.trim() !== '')
  const offBrand = text.filter((p) => p.font.name !== null && !ALLOWED_FONTS.has(p.font.name.toLowerCase()))
  if (!offBrand.length) return []
  const byFont = new Map<string, ParagraphInfo[]>()
  for (const p of offBrand) {
    const name = p.font.name ?? ''
    byFont.set(name, [...(byFont.get(name) ?? []), p])
  }
  const out: Finding[] = []
  for (const [font, ps] of byFont) {
    const first = ps[0]
    if (!first) continue
    // A document mostly in one other font needs its styles fixed, not hundreds of paragraphs.
    if (ps.length > Math.max(20, text.length / 2)) {
      out.push({
        id: `brand:font:${font}`, rule: 'brand/document-font', group: 'brand', severity: 'minor',
        title: `Most text is in ${font}`, detail: 'BwB uses Candara for headings and Aptos for text.',
        location: where(first), fix: { kind: 'add-styles' }, fixLabel: 'Restyle the document to BwB fonts',
      })
      continue
    }
    for (const p of ps) {
      const target = headingLevel(p) !== null ? 'Candara' : 'Aptos'
      out.push({
        id: `brand:font:${p.index}`, rule: 'brand/font', group: 'brand', severity: 'minor',
        title: `Text in ${font}`, detail: 'BwB uses Candara for headings and Aptos for text.',
        excerpt: excerpt(p.text, 0, 60), location: where(p),
        fix: { kind: 'paragraph-font', paragraph: p.index, expectedText: p.text, font: target }, fixLabel: `Change to ${target}`,
      })
    }
  }
  return out
}

function colourFindings(snapshot: DocSnapshot): Finding[] {
  const out: Finding[] = []
  const flagColour = (p: ParagraphInfo | undefined, colour: string, n: number | null): void => {
    const lime = colour === COLOURS.lime
    if (lime && p?.inTable) return
    if (!lime && ALLOWED_HEX.has(colour)) return
    out.push({
      id: `brand:colour:${n ?? 'x'}:${colour}`, rule: lime ? 'brand/lime-on-light' : 'brand/off-palette', group: 'brand', severity: 'minor',
      title: lime ? 'Lime text on a light background' : `Colour #${colour} is not in the BwB palette`,
      detail: lime ? 'Lime is for navy backgrounds only; it fails contrast on white.' : 'Use navy, charcoal or black for text; green only at 18 pt and above.',
      ...(p ? { excerpt: excerpt(p.text, 0, 60), location: where(p) } : {}),
    })
  }
  const highlight = (p: ParagraphInfo | undefined, n: number | null): void => {
    out.push({
      id: `brand:highlight:${n ?? 'x'}`, rule: 'brand/highlight', group: 'send', severity: 'minor',
      title: 'Highlighting left in', detail: 'Highlighting usually marks something still to do. BwB documents use callouts instead.',
      ...(p ? { excerpt: excerpt(p.text, 0, 60), location: where(p), fix: { kind: 'remove-highlight' as const, paragraph: p.index, expectedText: p.text }, fixLabel: 'Remove highlighting' } : {}),
    })
  }
  const small = (p: ParagraphInfo | undefined, pt: number, n: number | null): void => {
    out.push({
      id: `brand:small:${n ?? 'x'}`, rule: 'brand/small-text', group: 'brand', severity: 'minor',
      title: `Text at ${pt} pt`, detail: `Nothing in a BwB document is set below ${MIN_TEXT_PT} pt.`,
      ...(p ? { excerpt: excerpt(p.text, 0, 60), location: where(p) } : {}),
    })
  }

  if (snapshot.formatting) {
    for (const f of snapshot.formatting) {
      const p = f.paragraph === null ? undefined : snapshot.paragraphs[f.paragraph]
      f.colours.forEach((c) => flagColour(p, c, f.paragraph))
      if (f.highlights.length) highlight(p, f.paragraph)
      if (f.smallestHalfPoints !== null && f.smallestHalfPoints < MIN_TEXT_PT * 2) small(p, f.smallestHalfPoints / 2, f.paragraph)
    }
  } else {
    for (const p of snapshot.paragraphs) {
      if (!p.text.trim()) continue
      const c = hex(p.font.color)
      if (c) flagColour(p, c, p.index)
      if (p.font.highlight) highlight(p, p.index)
      if (p.font.size !== null && p.font.size < MIN_TEXT_PT) small(p, p.font.size, p.index)
    }
  }
  return out
}

export function brandFindings(snapshot: DocSnapshot): Finding[] {
  const findings = [...templateFindings(snapshot), ...fontFindings(snapshot), ...colourFindings(snapshot)]
  const seen = new Set<string>()
  return findings.filter((f) => (seen.has(f.id) ? false : (seen.add(f.id), true)))
}
