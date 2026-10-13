import { useState } from 'react'
import { LocateFixed, Wand2 } from 'lucide-react'
import type { CheckGroup, Finding, FixAction, Severity } from '../../check/types.ts'
import { Button } from '../components.tsx'

export const GROUPS: Array<{ id: CheckGroup; label: string }> = [
  { id: 'send', label: 'Before you send' },
  { id: 'structure', label: 'Structure and headings' },
  { id: 'tables', label: 'Tables and figures' },
  { id: 'brand', label: 'Brand and format' },
  { id: 'consistency', label: 'Consistency' },
  { id: 'writing', label: 'House style wording' },
]

export const SEVERITY_LABEL: Record<Severity, string> = { blocker: 'Blocker', major: 'Major', minor: 'Minor', copyedit: 'Copyedit' }

/** Fixes that change only what the finding names, so they can run in bulk without review. */
const BULK_SAFE = new Set<FixAction['kind']>(['replace', 'table-style', 'paragraph-font', 'remove-highlight', 'properties'])

export function bulkFixes(findings: Finding[]): Finding[] {
  return findings.filter((f) => !f.advisory && f.fix && BULK_SAFE.has(f.fix.kind))
}

const PAGE = 25

export interface FindingActions {
  busy: string | null
  onGoTo: (f: Finding) => void
  onFix: (f: Finding) => void
  onFixMany: (key: string, findings: Finding[]) => void
}

function FindingRow({ f, actions }: { f: Finding; actions: FindingActions }) {
  const working = actions.busy !== null
  return (
    <div className="finding">
      <div className="row" style={{ marginTop: 0, alignItems: 'flex-start', flexWrap: 'nowrap' }}>
        <span className={`badge ${f.advisory ? 'copyedit' : f.severity}`}>{f.advisory ? 'Suggestion' : SEVERITY_LABEL[f.severity]}</span>
        <span className="finding-title grow">{f.title}</span>
      </div>
      {f.detail && <div className="finding-detail">{f.detail}</div>}
      {f.excerpt && <div className="finding-excerpt">{f.excerpt}</div>}
      {(f.location || f.fix) && (
        <div className="row">
          {f.location && (
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

export function FindingGroup({ id, label, findings, actions }: { id: CheckGroup; label: string; findings: Finding[]; actions: FindingActions }) {
  const [shown, setShown] = useState(PAGE)
  const required = findings.filter((f) => !f.advisory)
  const urgent = required.some((f) => f.severity === 'blocker' || f.severity === 'major')
  const bulk = bulkFixes(findings)
  const worst = required[0]?.severity
  return (
    <details className="group" open={urgent || id === 'send'}>
      <summary>
        <span className="grow">{label}</span>
        {worst ? <span className={`badge ${worst}`}>{required.length}</span> : <span className="badge ok">{findings.length ? 'Suggestions' : 'Clear'}</span>}
      </summary>
      {bulk.length > 1 && (
        <div className="finding" style={{ background: 'var(--off-white)' }}>
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
