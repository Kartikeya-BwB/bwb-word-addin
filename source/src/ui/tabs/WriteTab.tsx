import type { CSSProperties } from 'react'
import { HEADING_STYLES, LIST_STYLES, TEXT_STYLES, UNNUMBERED_HEADING_STYLES } from '../../brand/styles.ts'
import type { ListStyle, ParagraphStyle } from '../../brand/styles.ts'
import { applyListStyle, applyParagraphStyle } from '../../word/styles.ts'
import { cleanSelectionWording } from '../../word/clean.ts'
import { Button, Section, Tile } from '../components.tsx'
import { useAction } from '../useAction.ts'

function previewStyle(s: ParagraphStyle): CSSProperties {
  const p = s.preview
  return {
    fontFamily: `'${p.font}', var(--font-body)`,
    fontSize: Math.min(p.sizePt * 1.05, 17),
    color: p.colour,
    fontWeight: p.bold ? 700 : 400,
    fontStyle: p.italic ? 'italic' : 'normal',
  }
}

export function WriteTab() {
  const { busy, run } = useAction()

  const apply = (s: ParagraphStyle): void =>
    void run(s.id, async () => {
      const r = await applyParagraphStyle(s)
      return `${s.label} applied${r.importedStyles ? '. BwB styles were added to this document' : ''}.`
    })
  const applyList = (l: ListStyle): void =>
    void run(l.id, async () => {
      const r = await applyListStyle(l)
      return `${l.label} list applied${r.importedStyles ? '. BwB styles were added to this document' : ''}.`
    })

  const styleTile = (s: ParagraphStyle, mark: string) => (
    <Tile key={s.id} mark={mark} label={s.label} hint={s.hint} preview={previewStyle(s)} onClick={() => apply(s)} title={`Apply ${s.label}`} />
  )

  return (
    <>
      <Section title="Headings" intro="Put the cursor in a paragraph, or select several, then choose a style. Headings 1 to 3 number themselves.">
        <div className="tiles">{HEADING_STYLES.map((s, i) => styleTile(s, `H${i + 1}`))}</div>
        <h3>Front and back matter</h3>
        <div className="tiles">{UNNUMBERED_HEADING_STYLES.map((s, i) => styleTile(s, `H${i + 1}`))}</div>
      </Section>

      <Section title="Text">
        <div className="tiles">{TEXT_STYLES.map((s) => styleTile(s, s.id === 'Normal' ? 'Aa' : s.id === 'BwBPullQuote' ? '“' : 'T'))}</div>
      </Section>

      <Section title="Lists" intro="Green squares for body lists. Use no more than two bullet styles on a page.">
        <div className="tiles">
          {LIST_STYLES.map((l) => (
            <Tile
              key={l.id}
              mark={l.marker}
              markClass={l.onNavy ? '' : 'light'}
              markStyle={{ color: l.markerColour, fontSize: l.numbered ? 13 : 18 }}
              label={l.label}
              hint={l.hint}
              disabled={busy !== null}
              onClick={() => applyList(l)}
              title={`Apply the ${l.label} list style`}
            />
          ))}
        </div>
      </Section>

      <Section title="House style in the selection" intro="Fixes wording that has one safe answer: UK spelling, dashes, dates, currency codes, per cent, the company name and double spaces.">
        <Button variant="navy" block busy={busy === 'clean'} onClick={() => void run('clean', cleanSelectionWording)}>
          Clean up the selected text
        </Button>
      </Section>
    </>
  )
}
