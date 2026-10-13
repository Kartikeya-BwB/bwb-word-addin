import { useCallback, useEffect, useState } from 'react'
import { ClipboardCopy, RefreshCw, ShieldCheck, Wand2 } from 'lucide-react'
import { qaNote, runChecks } from '../../check/engine.ts'
import type { CheckReport, Finding } from '../../check/types.ts'
import { SEVERITIES } from '../../check/types.ts'
import { plural } from '../../check/util.ts'
import { updateFields } from '../../word/fields.ts'
import { applyFix, applyFixes, goTo } from '../../word/fixes.ts'
import { insertAbbreviations } from '../../word/parts.ts'
import type { DocInfo } from '../../word/properties.ts'
import { hasWord } from '../../word/run.ts'
import { takeSnapshot } from '../../word/snapshot.ts'
import type { Sticker } from '../../word/stickers.ts'
import { STICKERS, currentSticker, setSticker } from '../../word/stickers.ts'
import { Button, Checkbox, Disclosure, Notice, Section, SelectField } from '../components.tsx'
import { useNotify } from '../toast.tsx'
import { useAction } from '../useAction.ts'
import { Checklist } from './Checklist.tsx'
import type { FindingActions } from './Findings.tsx'
import { FindingGroup, GROUPS, SEVERITY_LABEL, bulkFixes } from './Findings.tsx'

const ACRONYM_STATUS: Record<string, string> = {
  defined: 'Spelled out',
  'defined-late': 'Spelled out after first use',
  undefined: 'Never spelled out',
  common: 'Common, no need',
}

function StickerControl() {
  const { busy, run } = useAction()
  const [value, setValue] = useState<Sticker | ''>('')
  useEffect(() => {
    if (!hasWord()) return
    currentSticker().then(
      (current) => setValue(STICKERS.find((s) => s.toUpperCase() === (current ?? '').toUpperCase()) ?? ''),
      () => undefined,
    )
  }, [])
  return (
    <div className="row" style={{ alignItems: 'flex-end' }}>
      <div className="grow">
        <SelectField<Sticker | ''> label="Sticker in the page header" value={value} onChange={setValue}
          options={[{ value: '', label: 'None' }, ...STICKERS.map((s) => ({ value: s, label: s }))]} />
      </div>
      <Button style={{ marginBottom: 9 }} busy={busy === 'sticker'}
        onClick={() => void run('sticker', async () => {
          const n = await setSticker(value || null)
          return value ? `“${value}” sticker added to ${n} header${n === 1 ? '' : 's'}.` : 'Stickers removed.'
        })}>
        Apply
      </Button>
    </div>
  )
}

export function CheckTab() {
  const { busy, run } = useAction()
  const notify = useNotify()
  const [deep, setDeep] = useState(true)
  const [report, setReport] = useState<CheckReport | null>(null)
  const [info, setInfo] = useState<DocInfo | null>(null)

  const check = useCallback(async (): Promise<CheckReport> => {
    const snapshot = await takeSnapshot({ deepScan: deep })
    const next = runChecks(snapshot)
    setReport(next)
    setInfo(snapshot.info)
    return next
  }, [deep])

  const runCheck = (): void =>
    void run('check', async () => {
      const r = await check()
      const open = r.counts.blocker + r.counts.major + r.counts.minor + r.counts.copyedit
      const serious = r.counts.blocker + r.counts.major
      if (r.ready) return `Check complete: no blocker or major items${open ? `, ${open === 1 ? 'one smaller one' : `${open} smaller ones`}` : ''}.`
      return `Check complete: ${serious === 1 ? 'one item' : `${serious} items`} to resolve before QA.`
    })

  const actions: FindingActions = {
    busy,
    onGoTo: (f: Finding) => {
      const location = f.location
      if (location) void run(`go:${f.id}`, () => goTo(location))
    },
    onFix: (f: Finding) => {
      const fix = f.fix
      if (!fix) return
      void run(`fix:${f.id}`, async () => {
        await applyFix(fix)
        await check()
        return 'Fixed.'
      })
    },
    onFixMany: (key: string, findings: Finding[]) =>
      void run(key, async () => {
        const { applied, skipped } = await applyFixes(findings.flatMap((f) => (f.fix ? [f.fix] : [])))
        await check()
        return `${applied} fixed${skipped ? `; ${skipped} skipped because the text had changed` : ''}.`
      }),
  }

  const copyNote = async (): Promise<void> => {
    if (!report) return
    try {
      await navigator.clipboard.writeText(qaNote(report, info?.title ?? ''))
      notify('QA note copied. Paste it into the QA request.')
    } catch {
      notify('The clipboard is not available here.', 'error')
    }
  }

  const bulk = report ? bulkFixes(report.findings) : []

  return (
    <>
      <Section title="Pre-QA check" weight="act"
        intro="Checks the document against the BwB format, house style and QA checklist before peer review. Nothing changes until you choose Fix.">
        <Checkbox label="Also read fonts and colours run by run" checked={deep} onChange={setDeep}
          hint="Finds stray fonts, off-brand colours and text below 8 pt. Slower on very long documents." />
        <Button variant="primary" block busy={busy === 'check'} onClick={runCheck}>
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
            <Notice>No blocker or major items: ready for peer review and QA once the checklist is done.</Notice>
          ) : (
            <Notice kind="warn">Resolve the blocker and major items before the document goes to QA.</Notice>
          )}
          <p className="muted small" style={{ margin: '0 0 10px' }}>
            {[
              `${report.stats.words.toLocaleString('en-GB')} ${report.stats.words === 1 ? 'word' : 'words'}`,
              plural(report.stats.headings, 'heading'),
              plural(report.stats.tables, 'table'), plural(report.stats.pictures, 'picture'),
            ].join(' · ')}
          </p>
          {bulk.length > 1 && (
            <Button variant="navy" block busy={busy === 'fix-all'} onClick={() => actions.onFixMany('fix-all', bulk)}>
              <Wand2 size={14} aria-hidden /> Fix all {bulk.length} safe items
            </Button>
          )}
          <div style={{ marginTop: 10 }}>
            {GROUPS.map((g) => (
              <FindingGroup key={`${g.id}:${report.findings.length}`} id={g.id} label={g.label} actions={actions}
                findings={report.findings.filter((f) => f.group === g.id)} />
            ))}
          </div>

          {report.acronyms.length > 0 && (
            <Disclosure title={`Acronyms (${report.acronyms.length})`}>
              <table className="mini-table">
                <tbody>
                  {report.acronyms.map((a) => (
                    <tr key={a.acronym}>
                      <th scope="row">{a.acronym}</th>
                      <td>{a.definition ?? <span className="muted">{ACRONYM_STATUS[a.status] ?? a.status}</span>}</td>
                    </tr>
                  ))}
                </tbody>
              </table>
              <Button small block busy={busy === 'aa'} onClick={() => void run('aa', insertAbbreviations)}>Insert an abbreviations table at the cursor</Button>
            </Disclosure>
          )}
        </>
      )}

      <Section title="Checklist" intro="The QA policy’s author self-check, plus checks for the document type and funder set on the Start tab.">
        <Checklist report={report} info={info} />
      </Section>

      <Section title="Send to QA">
        <StickerControl />
        <div className="row">
          <Button busy={busy === 'fields'}
            onClick={() => void run('fields', async () => {
              const n = await updateFields()
              return `${n} field${n === 1 ? '' : 's'} updated: contents, captions and cross-references.`
            })}>
            <RefreshCw size={13} aria-hidden /> Update fields
          </Button>
          <Button disabled={!report} onClick={() => void copyNote()} title={report ? undefined : 'Run the check first'}>
            <ClipboardCopy size={13} aria-hidden /> Copy QA note
          </Button>
        </div>
      </Section>
    </>
  )
}
