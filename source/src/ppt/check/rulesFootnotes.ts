// Footnotes: every superscript marker on a slide has a note in the slide's footnote block, and every note has a
// marker. Markers are read only where PowerPoint reports superscript (API 1.8).

import { FOOTNOTES_TAG, missingFootnotes, parseFootnotes, unusedFootnotes } from '../slides/footnotes.ts'
import { hasTag } from './shapes.ts'
import type { DeckFinding, DeckSnapshot } from './types.ts'
import { excerpt, onShape } from './util.ts'

export function footnoteFindings(snapshot: DeckSnapshot): DeckFinding[] {
  if (!snapshot.can.markers) return []
  const out: DeckFinding[] = []
  for (const slide of snapshot.slides) {
    const block = slide.shapes.find((s) => hasTag(s, FOOTNOTES_TAG))
    const notes = block ? parseFootnotes(block.text ?? '') : []
    const markers = slide.shapes.flatMap((shape) => shape.markers.map((marker) => ({ shape, marker })))
    const n = slide.index + 1
    for (const number of missingFootnotes(markers.map((m) => m.marker.number), notes)) {
      const first = markers.find((m) => m.marker.number === number)
      if (!first) continue
      const { shape, marker } = first
      out.push({
        id: `slides:footnote-missing:${slide.id}:${number}`, rule: 'slides/footnote-missing', group: 'slides', severity: 'minor',
        title: `Footnote marker ${number} has no footnote`,
        detail: 'Add the note above the source line (Footnotes on the Slides tab), or remove the marker.',
        excerpt: excerpt(shape.text ?? '', marker.start, marker.start + marker.length), slide: n,
        location: onShape(slide, shape, marker.start, marker.start + marker.length),
      })
    }
    if (!block) continue
    for (const note of unusedFootnotes(markers.map((m) => m.marker.number), notes)) {
      out.push({
        id: `slides:footnote-unused:${slide.id}:${note.number}`, rule: 'slides/footnote-unused', group: 'slides', severity: 'minor',
        title: `Footnote ${note.number} has no marker`, detail: 'Nothing on the slide refers to it. Add its marker in the text, or remove the note.',
        excerpt: excerpt(`${note.number} ${note.text}`, 0, 60, 0), slide: n, location: onShape(slide, block),
        fix: { kind: 'remove-unused-footnotes', slideId: slide.id }, fixLabel: 'Remove the notes with no marker',
      })
    }
  }
  return out
}
