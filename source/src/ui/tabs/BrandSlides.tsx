// The parts of the Brand tab that differ for slides: colour actions on shapes, the slide type scale, the grid and
// what goes where on a slide.

import { PaintBucket, Type } from 'lucide-react'
import type { ColourName } from '../../brand/tokens.ts'
import { SLIDE_TYPE } from '../../brand/slideGrid.ts'
import type { FillChoice, TextColour } from '../../ppt/format.ts'
import { applyFill, applyTextColour } from '../../ppt/format.ts'
import { Button, Section } from '../components.tsx'
import { SlideGridDiagram } from '../library/SlideGridDiagram.tsx'
import { useAction } from '../useAction.ts'
import { RuleTable } from './RuleTable.tsx'

/** What each palette colour can do to the selected shapes. Lime and the greys have no action: they are accents and lines. */
const SHAPE_ACTIONS: Partial<Record<ColourName, { fill?: FillChoice; text?: TextColour }>> = {
  navy: { fill: 'navy', text: 'navy' },
  green: { fill: 'green' },
  offWhite: { fill: 'offWhite' },
  charcoal: { text: 'charcoal' },
  white: { text: 'white' },
}

/** The buttons beside a swatch on the PowerPoint Brand tab. */
export function SlideColourActions({ colour, name }: { colour: ColourName; name: string }) {
  const { busy, run } = useAction()
  const action = SHAPE_ACTIONS[colour]
  if (!action) return null
  const { fill, text } = action
  return (
    <>
      {fill && (
        <Button small title={`Fill the selected shapes ${name.toLowerCase()}`} aria-label={`Fill the selected shapes ${name.toLowerCase()}`}
          busy={busy === 'fill'} onClick={() => void run('fill', () => applyFill(fill))}>
          <PaintBucket size={13} aria-hidden />
        </Button>
      )}
      {text && (
        <Button small title={`Colour the text of the selected shapes ${name.toLowerCase()}`} aria-label={`Colour the selected text ${name.toLowerCase()}`}
          busy={busy === 'text'} onClick={() => void run('text', () => applyTextColour(text))}>
          <Type size={13} aria-hidden />
        </Button>
      )}
    </>
  )
}

const TYPE_ROWS: Array<[string, { font: string; size: number; bold: boolean }, string]> = [
  ['Action title', SLIDE_TYPE.title, 'Navy, one or two lines, a full sentence that states the point'],
  ['Subtitle', SLIDE_TYPE.subtitle, 'Charcoal, under the title when the slide needs one'],
  ['Body', SLIDE_TYPE.body, `Black; never below ${SLIDE_TYPE.minBody} pt`],
  ['Small text', SLIDE_TYPE.small, 'Labels, table text and annotations'],
  ['Source', SLIDE_TYPE.source, `Charcoal, above the footer: “Source: ...”. Nothing on a slide below ${SLIDE_TYPE.minAny} pt`],
]

const SLIDE_FURNITURE: Array<[string, string]> = [
  ['Size', '16:9 widescreen, 33.87 × 19.05 cm'],
  ['Title', 'An action title in the title placeholder at the top left, never in a text box'],
  ['Content', 'Inside the margins and on the 12 columns; line edges up with Arrange'],
  ['Source', 'A source line above the footer for every chart, table and figure'],
  ['Footer', 'The BwB logo at the bottom left and the slide number at the bottom right'],
  ['Stickers', 'Draft, For discussion and the rest at the top right; remove them before you send'],
  ['Ending', 'The BwB closing slide from the slide library'],
]

const SLIDE_ICON_RULES: Array<[string, string]> = [
  ['On white', 'Green or navy; charcoal for notes'],
  ['On navy panels', 'White, or lime for the one accent on the panel'],
  ['Sizes', '24 pt beside text, 48 pt as tile markers, 96 pt on a section divider or panel'],
  ['Use', 'Each icon names a topic, marks a callout or labels a tile. Give every icon alt text'],
]

export function SlideBrandSections() {
  return (
    <>
      <Section title="Type on slides" intro="Candara for titles, Aptos for everything else.">
        <table className="mini-table">
          <tbody>
            {TYPE_ROWS.map(([element, t, detail]) => (
              <tr key={element}>
                <th scope="row">{element}</th>
                <td>
                  <span style={{ fontFamily: `'${t.font}', var(--font-body)` }}>{t.font} {t.size} pt{t.bold ? ' bold' : ''}</span>
                  <div className="tile-hint">{detail}</div>
                </td>
              </tr>
            ))}
          </tbody>
        </table>
      </Section>

      <Section id="grid" title="The slide grid">
        <SlideGridDiagram />
      </Section>

      <Section title="Slide furniture">
        <RuleTable rows={SLIDE_FURNITURE} />
      </Section>

      <Section title="Icons">
        <RuleTable rows={SLIDE_ICON_RULES} />
      </Section>
    </>
  )
}
