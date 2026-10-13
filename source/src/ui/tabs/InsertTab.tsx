import {
  Award, Briefcase, CaseUpper, Grid2x2, Hash, Info, LayoutGrid, Lightbulb, Quote, Scale,
  SeparatorHorizontal, SquarePlus, TableOfContents, TableProperties,
} from 'lucide-react'
import type { LucideIcon } from 'lucide-react'
import type { LibraryBlock } from '../../brand/library.ts'
import { LIBRARY } from '../../brand/library.ts'
import { insertBlocks } from '../../word/insert.ts'
import { HARVEY_BALLS, insertHarveyBall, insertPageBreak } from '../../word/marks.ts'
import { insertAbbreviations, insertContents, insertDisclaimer, insertEuAcknowledgement } from '../../word/parts.ts'
import { ChartBuilder } from '../charts/ChartBuilder.tsx'
import { Disclosure, Section, Tile } from '../components.tsx'
import { useAction } from '../useAction.ts'
import { IconBrowser } from '../library/IconBrowser.tsx'
import { ImagePicker } from '../library/ImagePicker.tsx'
import { LogoPanel } from '../library/LogoPanel.tsx'
import { CalloutForm, KeyFiguresForm, PullQuoteForm, SnapshotForm } from './PanelForms.tsx'
import { CaptionForm, DataTableForm, FormatTable, SourceAndFootnote } from './TableForms.tsx'

// One icon per library block (src/brand/library.ts).
const BLOCK_ICONS: Record<string, LucideIcon> = {
  callout: Lightbulb, note: Info, stats: Hash, snapshot: LayoutGrid, quote: Quote, case: Briefcase, grid: Grid2x2,
  details: TableProperties,
}

interface DocumentPart {
  id: string
  label: string
  hint: string
  Icon: LucideIcon
  action: () => Promise<string>
}

const DOCUMENT_PARTS: DocumentPart[] = [
  { id: 'toc', label: 'Contents', hint: 'Headings 1 and 2, then a page break', Icon: TableOfContents, action: insertContents },
  { id: 'disclaimer', label: 'Disclaimer', hint: 'The standard BwB disclaimer on its own page at the end', Icon: Scale, action: insertDisclaimer },
  { id: 'eu', label: 'EU acknowledgement', hint: 'Funding statement and disclaimer for EU-funded work', Icon: Award, action: insertEuAcknowledgement },
  { id: 'aa', label: 'Abbreviations table', hint: 'Every acronym in the document with the meaning it gives', Icon: CaseUpper, action: insertAbbreviations },
  {
    id: 'break', label: 'Page break', hint: 'Start the next part on a new page', Icon: SeparatorHorizontal,
    action: async () => {
      await insertPageBreak()
      return 'Page break inserted.'
    },
  },
]


const mark = (Icon: LucideIcon) => <Icon size={18} strokeWidth={1.8} />

export function InsertTab() {
  const { busy, run } = useAction()
  const working = busy !== null

  const insertBlock = (b: LibraryBlock): void =>
    void run(b.id, async () => {
      const { warnings } = await insertBlocks(b.blocks())
      const done = `${b.label} inserted. Type over the grey placeholders.`
      return warnings.length ? `${done} ${warnings.join(' ')}` : done
    })

  return (
    <>
      <Section title="Building blocks" weight="act"
        intro="One click puts a block below the cursor with grey placeholders to type over. The Check tab lists any left unfilled.">
        <div className="tiles">
          {LIBRARY.map((b) => (
            <Tile key={b.id} mark={mark(BLOCK_ICONS[b.id] ?? SquarePlus)} label={b.label} hint={b.hint} disabled={working} onClick={() => insertBlock(b)} title={`Insert a ${b.label.toLowerCase()}`} />
          ))}
        </div>
      </Section>

      <Section title="Write it first" intro="Fill in the text here, then insert it in one go.">
        <Disclosure title="Callout or note"><CalloutForm /></Disclosure>
        <Disclosure title="Key figures"><KeyFiguresForm /></Disclosure>
        <Disclosure title="Pull quote"><PullQuoteForm /></Disclosure>
        <Disclosure title="Snapshot"><SnapshotForm /></Disclosure>
      </Section>

      <Section title="Tables and captions" intro="Navy header row repeated on each page, figures right-aligned, caption below.">
        <Disclosure title="New data table"><DataTableForm /></Disclosure>
        <Disclosure title="Format an existing table"><FormatTable /></Disclosure>
        <Disclosure title="Caption for a figure or table"><CaptionForm /></Disclosure>
        <Disclosure title="Source line and footnote"><SourceAndFootnote /></Disclosure>
      </Section>

      <Section title="Document parts">
        <div className="tiles">
          {DOCUMENT_PARTS.map((p) => (
            <Tile key={p.id} mark={mark(p.Icon)} label={busy === p.id ? 'Working…' : p.label} hint={p.hint} disabled={working}
              onClick={() => void run(p.id, p.action)} />
          ))}
        </div>
      </Section>

      <Section title="Harvey balls" intro="Score options against criteria in a table. Inserted at the cursor in navy.">
        <div className="row">
          {HARVEY_BALLS.map((h) => (
            <button key={h.glyph} type="button" className="btn harvey" title={h.label} aria-label={`Harvey ball: ${h.label}`} disabled={working}
              onClick={() => void run(`harvey-${h.glyph}`, async () => {
                await insertHarveyBall(h.glyph)
              })}>
              {h.glyph}
            </button>
          ))}
        </div>
      </Section>

      <Section id="charts" title="Charts"
        intro="Column, bar, line, waterfall, Mekko, Gantt, donut and Harvey ball tables in the BwB colours, inserted as a sharp picture with a numbered caption.">
        <ChartBuilder />
      </Section>

      <Section id="photos" title="Photos" intro="The BwB photo library: sectors, themes, markets, and the team. Double-click a photo to insert it.">
        <ImagePicker />
      </Section>

      <Section id="icons" title="Icons, flags and maps"
        intro="Lucide (the BwB set), the BwB sectors, Tabler, flags and maps in the brand colours and sizes. Double-click to insert.">
        <IconBrowser />
      </Section>

      <Section title="Logo">
        <LogoPanel />
      </Section>
    </>
  )
}
