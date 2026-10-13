import {
  ChevronsRight, CircleAlert, CircleCheck, CircleX, Grid2x2, Hash, Lightbulb, ListOrdered, Milestone, Quote, SquarePlus, Table, TextAlignStart,
} from 'lucide-react'
import type { LucideIcon } from 'lucide-react'
import { ELEMENTS } from '../../ppt/elements.ts'
import { addReviewNote, insertHarveyBall, insertRagDot, insertStatusMark, removeAllNotes } from '../../ppt/marks.ts'
import { hasPowerPoint, supportsPpt, updateMessage } from '../../ppt/run.ts'
import type { RagColour } from '../../ppt/specs.ts'
import type { Quarters, StatusKind } from '../../ppt/svg.ts'
import { HARVEY_LABELS, STATUS_MARKS } from '../../ppt/svg.ts'
import { HARVEY_BALLS } from '../../word/marks.ts'
import { Button, Notice, Section, Tile } from '../components.tsx'
import { useAction } from '../useAction.ts'
import './ppt.css'

const ICONS: Record<string, LucideIcon> = {
  callout: Lightbulb, 'key-figures': Hash, quote: Quote, process: ChevronsRight, matrix: Grid2x2, timeline: Milestone,
  'step-marker': ListOrdered, source: TextAlignStart, table: Table,
}

const STATUS: Array<{ kind: StatusKind; Icon: LucideIcon; className: string }> = [
  { kind: 'onTrack', Icon: CircleCheck, className: 'status-on-track' },
  { kind: 'atRisk', Icon: CircleAlert, className: 'status-other' },
  { kind: 'offTrack', Icon: CircleX, className: 'status-other' },
]

const RAG: Array<{ colour: RagColour; label: string }> = [
  { colour: 'green', label: 'Green dot' },
  { colour: 'amber', label: 'Amber dot, for trackers only' },
  { colour: 'red', label: 'Red dot, for trackers only' },
]

const QUARTERS: Quarters[] = [0, 1, 2, 3, 4]

/** BwB slide elements, Harvey balls, status marks and review notes for the current slide. Stickers are on the Slides tab. */
export function ElementsPanel() {
  const { busy, run } = useAction()
  const live = hasPowerPoint()
  const why = (version: string, feature: string): string | null => (live && !supportsPpt(version) ? updateMessage(feature) : null)
  const blocked = why('1.5', 'Inserting slide elements')
  const off = busy !== null || blocked !== null

  return (
    <>
      {blocked && <Notice kind="warn">{blocked}</Notice>}
      <Section title="Slide elements" weight="act" intro="One click puts a BwB element in the content area of the current slide, ready to type over.">
        <div className="tiles">
          {ELEMENTS.map((e) => {
            const Icon = ICONS[e.id] ?? SquarePlus
            const reason = why(e.version, 'This element')
            return (
              <Tile key={e.id} mark={<Icon size={18} strokeWidth={1.8} />} label={busy === e.id ? 'Working…' : e.label} hint={reason ?? e.hint}
                disabled={off || reason !== null} title={reason ?? `Insert ${e.label.toLowerCase()}`} onClick={() => void run(e.id, e.insert)} />
            )
          })}
        </div>
      </Section>

      <Section title="Harvey balls" intro="In navy, as graphics. Click again to add the next one beside it.">
        <div className="mark-row">
          {QUARTERS.map((q) => (
            <button key={q} type="button" className="btn harvey" title={HARVEY_LABELS[q]} aria-label={HARVEY_LABELS[q]} disabled={off}
              onClick={() => void run(`harvey-${q}`, () => insertHarveyBall(q))}>
              {HARVEY_BALLS[q]?.glyph}
            </button>
          ))}
        </div>
      </Section>

      <Section title="Status marks" intro="Ticks and crosses in BwB colours. Red, amber and green dots are for trackers only.">
        <div className="mark-row">
          {STATUS.map(({ kind, Icon, className }) => (
            <button key={kind} type="button" className={`btn ${className}`} title={STATUS_MARKS[kind].label} aria-label={STATUS_MARKS[kind].label} disabled={off}
              onClick={() => void run(kind, () => insertStatusMark(kind))}>
              <Icon size={20} strokeWidth={2} aria-hidden />
            </button>
          ))}
          {RAG.map(({ colour, label }) => (
            <button key={colour} type="button" className="btn" title={label} aria-label={label} disabled={off}
              onClick={() => void run(`rag-${colour}`, () => insertRagDot(colour))}>
              <span className={`rag-dot ${colour}`} aria-hidden />
            </button>
          ))}
        </div>
      </Section>

      <Section title="Review notes" intro="Dated navy notes for reviewers. Remove them all in one click before you send the deck.">
        <div className="row">
          <Button variant="navy" busy={busy === 'note'} disabled={off} onClick={() => void run('note', () => addReviewNote())}>Add a note</Button>
          <Button disabled={off} onClick={() => void run('notes-off', removeAllNotes)}>Remove all notes</Button>
        </div>
      </Section>
    </>
  )
}
