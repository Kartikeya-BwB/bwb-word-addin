import { useCallback, useState } from 'react'
import { ClipboardCopy, LocateFixed, RefreshCw, ShieldCheck, Wand2 } from 'lucide-react'
import { SEVERITIES } from '../../check/types.ts'
import { deckBulkFixes, deckQaNote, runDeckCheck } from '../../ppt/check/engine.ts'
import { applyDeckFix, applyDeckFixes, goToLocation } from '../../ppt/check/fixes.ts'
import { slideTitle } from '../../ppt/check/shapes.ts'
import { takeDeckSnapshot } from '../../ppt/check/snapshot.ts'
import type { DeckFinding, DeckGroup, DeckReport } from '../../ppt/check/types.ts'
import { NEEDS } from '../../ppt/slides/office.ts'
import { supportsPpt, unavailableReason } from '../../ppt/run.ts'
import { Button, Notice, Section } from '../components.tsx'
import { SEVERITY_LABEL } from '../tabs/Findings.tsx'
import { useNotify } from '../toast.tsx'
import { useAction } from '../useAction.ts'
import './slides.css'

export const DECK_GROUPS: Array<{ id: DeckGroup; label: string }> = [
  { id: 'send', label: 'Before you send' },
  { id: 'brand', label: 'Brand' },
  { id: 'layout', label: 'Layout' },
  { id: 'wording', label: 'Wording' },
  { id: 'slides', label: 'Slides' },
]

const PAGE = 25

interface Actions {
  busy: string | null
  canGo: boolean
  onGoTo: (f: DeckFinding) => void
  onFix: (f: DeckFinding) => void
  onFixMany: (key: string, findings: DeckFinding[]) => void
}

function FindingRow({ f, actions }: { f: DeckFinding; actions: Actions }) {
  const working = actions.busy !== null
  return (
    <div className="finding">
      <div className="row" style={{ marginTop: 0, alignItems: 'flex-start', flexWrap: 'nowrap' }}>
        <span className={`badge ${f.advisory ? 'copyedit' : f.severity}`}>{f.advisory ? 'Suggestion' : SEVERITY_LABEL[f.severity]}</span>
        <span className="finding-title grow">{f.title}</span>
        {f.slide !== undefined && <span className="finding-slide">Slide {f.slide}</span>}
      </div>
      {f.detail && <div className="finding-detail">{f.detail}</div>}
      {f.excerpt && <div className="finding-excerpt">{f.excerpt}</div>}
      {((f.location && actions.canGo) || f.fix) && (
        <div className="row">
          {f.location && actions.canGo && (
            <Button small disabled={working} busy={actions.busy === `go:${f.id}`} onClick={() => actions.onGoTo(f)}>
              <LocateFixed size={13} aria-hidden /> Go to
            </Button>
          )}
          {f.fix && (
            <Button small variant="navy" disabled={working} busy={actions.busy === `fix:${f.id}`} onClick={() => actions.onFix(f)}>
              <Wand2 size={13} aria-hidden /> {f.fixLabel ?? 'Fix'}
            </Button>
          )}
        </div>
      )}
    </div>
  )
}

function FindingGroup({ id, label, findings, actions }: { id: DeckGroup; label: string; findings: DeckFinding[]; actions: Actions }) {
  const [shown, setShown] = useState(PAGE)
  const required = findings.filter((f) => !f.advisory)
  const urgent = required.some((f) => f.severity === 'blocker' || f.severity === 'major')
  const bulk = deckBulkFixes(findings)
  const worst = required[0]?.severity
  return (
    <details className="group" open={urgent || id === 'send'}>
      <summary>
        <span className="grow">{label}</span>
        {worst ? <span className={`badge ${worst}`}>{required.length}</span> : <span className="badge ok">{findings.length ? 'Suggestions' : 'Clear'}</span>}
      </summary>
      {bulk.length > 1 && (
        <div className="finding fix-group">
          <Button small variant="navy" disabled={actions.busy !== null} busy={actions.busy === `group:${id}`} onClick={() => actions.onFixMany(`group:${id}`, bulk)}>
            <Wand2 size={13} aria-hidden /> Fix all {bulk.length} in this group
          </Button>
        </div>
      )}
      {findings.length === 0 && <div className="finding muted small">Nothing to fix.</div>}
      {findings.slice(0, shown).map((f) => <FindingRow key={f.id} f={f} actions={actions} />)}
      {findings.length > shown && (
        <div className="finding">
          <Button small variant="link" onClick={() => setShown((n) => n + PAGE * 4)}>Show {Math.min(PAGE * 4, findings.length - shown)} more of {findings.length - shown}</Button>
        </div>
      )}
    </details>
  )
}

const plural = (n: number, one: string, many = `${one}s`): string => `${n} ${n === 1 ? one : many}`

/** The deck check: the PowerPoint counterpart of the Word Pre-QA check. */
export function DeckCheckTab() {
  const { busy, run } = useAction()
  const notify = useNotify()
  const [report, setReport] = useState<DeckReport | null>(null)
  const [deckTitle, setDeckTitle] = useState('')

  const check = useCallback(async (): Promise<DeckReport> => {
    const snapshot = await takeDeckSnapshot()
    const next = runDeckCheck(snapshot)
    setReport(next)
    const first = snapshot.slides[0]
    setDeckTitle(first ? slideTitle(first)?.text ?? '' : '')
    return next
  }, [])

  const runCheck = (): void =>
    void run('check', async () => {
      const r = await check()
      const open = r.counts.blocker + r.counts.major + r.counts.minor + r.counts.copyedit
      const serious = r.counts.blocker + r.counts.major
      if (r.ready) return `Check complete: no blocker or major items${open ? `, ${open === 1 ? 'one smaller one' : `${open} smaller ones`}` : ''}.`
      return `Check complete: ${serious === 1 ? 'one item' : `${serious} items`} to resolve before QA.`
    })

  const actions: Actions = {
    busy,
    canGo: supportsPpt(NEEDS.current),
    onGoTo: (f) => {
      const location = f.location
      if (location) void run(`go:${f.id}`, () => goToLocation(location))
    },
    onFix: (f) => {
      const fix = f.fix
      if (!fix) return
      void run(`fix:${f.id}`, async () => {
        await applyDeckFix(fix)
        await check()
        return 'Fixed.'
      })
    },
    onFixMany: (key, findings) =>
      void run(key, async () => {
        const { applied, skipped } = await applyDeckFixes(findings.flatMap((f) => (f.fix ? [f.fix] : [])))
        await check()
        return `${applied} fixed${skipped ? `; ${skipped} skipped because the slide had changed` : ''}.`
      }),
  }

  const copyNote = async (): Promise<void> => {
    if (!report) return
    try {
      await navigator.clipboard.writeText(deckQaNote(report, deckTitle))
      notify('QA note copied. Paste it into the QA request.')
    } catch {
      notify('The clipboard is not available here.', 'error')
    }
  }

  const tooOld = unavailableReason(NEEDS.read, 'The deck check')
  const bulk = report ? deckBulkFixes(report.findings) : []
  return (
    <>
      <Section title="Deck check" weight="act"
        intro="Checks every slide against the BwB slide format and house style before the deck goes to QA. Nothing changes until you choose Fix.">
        {tooOld && <Notice kind="warn">{tooOld}</Notice>}
        <Button variant="primary" block disabled={Boolean(tooOld)} busy={busy === 'check'} onClick={runCheck}>
          {report ? <><RefreshCw size={14} aria-hidden /> Run the check again</> : <><ShieldCheck size={14} aria-hidden /> Run the check</>}
        </Button>
      </Section>

      {report && (
        <>
          <div className="summary" role="group" aria-label="Open items by severity">
            {SEVERITIES.map((s) => (
              <div key={s}>
                <strong className={report.counts[s] ? `sev-${s}` : ''}>{report.counts[s]}</strong>
                <span>{SEVERITY_LABEL[s]}</span>
              </div>
            ))}
          </div>
          {report.ready ? (
            <Notice>No blocker or major items: ready for peer review and QA.</Notice>
          ) : (
            <Notice kind="warn">Resolve the blocker and major items before the deck goes to QA.</Notice>
          )}
          {report.limits.map((l) => <p key={l} className="field-hint deck-limit">{l}</p>)}
          <p className="muted small" style={{ margin: '0 0 10px' }}>
            {[
              plural(report.stats.slides, 'slide'),
              `${report.stats.words.toLocaleString('en-GB')} ${report.stats.words === 1 ? 'word' : 'words'}`,
              plural(report.stats.pictures, 'picture'),
            ].join(' · ')}
          </p>
          {bulk.length > 1 && (
            <Button variant="navy" block busy={busy === 'fix-all'} onClick={() => actions.onFixMany('fix-all', bulk)}>
              <Wand2 size={14} aria-hidden /> Fix all {bulk.length} safe items
            </Button>
          )}
          <div style={{ marginTop: 10 }}>
            {DECK_GROUPS.map((g) => (
              <FindingGroup key={`${g.id}:${report.findings.length}`} id={g.id} label={g.label} actions={actions}
                findings={report.findings.filter((f) => f.group === g.id)} />
            ))}
          </div>
        </>
      )}

      <Section title="Send to QA" intro="Paste the note into the QA request, so the reviewer sees what the check found.">
        <Button disabled={!report} onClick={() => void copyNote()} title={report ? undefined : 'Run the check first'}>
          <ClipboardCopy size={13} aria-hidden /> Copy QA note
        </Button>
      </Section>
    </>
  )
}
