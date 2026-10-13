import { Copy, Paintbrush } from 'lucide-react'
import { ICON_SIZES_CM } from '../../brand/icons.ts'
import type { ColourName } from '../../brand/tokens.ts'
import { PALETTE, TYPE_SCALE, WRITING_RULES } from '../../brand/tokens.ts'
import { colourSelection } from '../../word/marks.ts'
import { Button, Section } from '../components.tsx'
import { useNotify } from '../toast.tsx'
import { useAction } from '../useAction.ts'

/** Colours BwB allows for text on a white page, which the swatches can apply to the selection. */
const TEXT_COLOURS = new Set<ColourName>(['navy', 'charcoal', 'black'])

const ICON_RULES: Array<[string, string]> = [
  ['On white', 'Green or navy; charcoal for notes'],
  ['On navy panels', 'White, or lime for the one accent on the panel'],
  ['Sizes', `${ICON_SIZES_CM.inline} cm in text, ${ICON_SIZES_CM.callout} cm in callout headers, ${ICON_SIZES_CM.marker} cm as tile markers, ${ICON_SIZES_CM.panel} cm on panels`],
  ['Use', 'Each icon names a topic, marks a callout or labels a tile. Inline with text, never floating, always with alt text'],
]

const FURNITURE: Array<[string, string]> = [
  ['Page', 'A4 portrait, 2.2 cm margins'],
  ['Cover', 'Photograph above a navy panel: title in Candara 32 pt white, tagline, lime rule, “Bankers without Boundaries · D Month YYYY”'],
  ['Contents', 'A real contents field (levels 1 and 2) when there are five or more headings'],
  ['Headings', 'Headings 1 to 3 number themselves (1., 1.1, 1.1.1). Never type the numbers. Front and back matter are not numbered'],
  ['Footer', '“Page X of Y” bottom right, Aptos 8 pt charcoal, hidden on the cover and dividers'],
  ['Tables', 'Navy header row repeated on each page, light grey borders, figures right-aligned, total row off-white and bold'],
  ['Captions', 'Below tables and figures: “Table 1. Title. Source: ...”, 9 pt italic charcoal'],
  ['Ending', 'The standard disclaimer on its own page (FRN 953637)'],
]

export function BrandTab() {
  const { busy, run } = useAction()
  const notify = useNotify()

  const copy = async (text: string): Promise<void> => {
    try {
      await navigator.clipboard.writeText(text)
      notify(`${text} copied.`)
    } catch {
      notify('The clipboard is not available here.', 'error')
    }
  }

  return (
    <>
      <Section title="Colours" intro="Lime only on navy. Green text only at 18 pt and above. Body text is black.">
        <div className="tiles">
          {PALETTE.map((c) => (
            <div key={c.key} className="swatch-row">
              <span className="swatch" style={{ background: `#${c.hex}` }} aria-hidden />
              <div className="grow">
                <div className="tile-label">{c.name} <span className="mono muted">#{c.hex}</span></div>
                <div className="tile-hint">{c.use}</div>
                <div className="tile-hint"><strong>Text:</strong> {c.textRule}</div>
              </div>
              <div className="swatch-actions">
                {TEXT_COLOURS.has(c.key) && (
                  <Button small title={`Colour the selected text ${c.name.toLowerCase()}`} aria-label={`Colour the selected text ${c.name.toLowerCase()}`}
                    busy={busy === c.key} onClick={() => void run(c.key, async () => {
                      await colourSelection(c.hex)
                      return `Selected text set to ${c.name.toLowerCase()}.`
                    })}>
                    <Paintbrush size={13} aria-hidden />
                  </Button>
                )}
                <Button small title={`Copy #${c.hex}`} aria-label={`Copy the ${c.name} hex code`} onClick={() => void copy(`#${c.hex}`)}>
                  <Copy size={13} aria-hidden />
                </Button>
              </div>
            </div>
          ))}
        </div>
      </Section>

      <Section title="Type" intro="Candara for headings, Aptos for everything else. Nothing below 8 pt.">
        <table className="mini-table">
          <tbody>
            {TYPE_SCALE.map((t) => (
              <tr key={t.element}>
                <th scope="row">{t.element}</th>
                <td>
                  <span style={{ fontFamily: `'${t.font}', var(--font-body)` }}>{t.font} {t.sizePt} pt</span>
                  <div className="tile-hint">{t.detail}</div>
                </td>
              </tr>
            ))}
          </tbody>
        </table>
      </Section>

      <Section title="House style" intro="The Check tab finds these, and Clean up on the Write tab fixes the ones with one right answer.">
        <table className="mini-table">
          <tbody>
            {WRITING_RULES.map((w) => (
              <tr key={w.rule}>
                <th scope="row">{w.rule}</th>
                <td>{w.example}</td>
              </tr>
            ))}
          </tbody>
        </table>
      </Section>

      <Section title="Document furniture">
        <table className="mini-table">
          <tbody>
            {FURNITURE.map(([k, v]) => (
              <tr key={k}>
                <th scope="row">{k}</th>
                <td>{v}</td>
              </tr>
            ))}
          </tbody>
        </table>
      </Section>

      <Section title="Icons">
        <table className="mini-table">
          <tbody>
            {ICON_RULES.map(([k, v]) => (
              <tr key={k}>
                <th scope="row">{k}</th>
                <td>{v}</td>
              </tr>
            ))}
          </tbody>
        </table>
      </Section>

      <p className="muted small">
        Source: the BwB document format (v1.1.0) and the BwB Style Guide. Where a client template applies, it overrides BwB house style.
      </p>
    </>
  )
}
