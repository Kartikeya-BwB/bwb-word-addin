import { useCallback, useEffect, useState } from 'react'
import { Copy } from 'lucide-react'
import { DOC_TYPES } from '../../brand/docTypes.ts'
import { fileNameIssues, suggestFileName } from '../../check/fileName.ts'
import type { DocInfo, FunderFrame } from '../../word/properties.ts'
import { FUNDER_FRAMES, documentFileName, readDocInfo, writeDocInfo } from '../../word/properties.ts'
import { addBwbStyles, hasBwbStyles, restyleBuiltInStyles } from '../../word/styles.ts'
import { hasWord, supports } from '../../word/run.ts'
import { Button, Checkbox, Notice, SelectField, TextField } from '../components.tsx'
import { useNotify } from '../toast.tsx'
import { useAction } from '../useAction.ts'

const EMPTY: DocInfo = {
  title: '', subject: '', author: '', company: '', docType: '', project: '', deliverable: '', version: '', funder: '', clientTemplate: false,
}

export function ThisDocument() {
  const { busy, run } = useAction()
  const notify = useNotify()
  const [info, setInfo] = useState<DocInfo>(EMPTY)
  const [styles, setStyles] = useState<boolean | null>(null)
  const [loaded, setLoaded] = useState(false)
  const fileName = hasWord() ? documentFileName() : null

  const refresh = useCallback(async () => {
    if (!hasWord()) return
    try {
      setInfo(await readDocInfo())
      setStyles(await hasBwbStyles())
      setLoaded(true)
    } catch {
      setLoaded(false)
    }
  }, [])

  useEffect(() => {
    void refresh()
  }, [refresh])

  const set = <K extends keyof DocInfo>(key: K) => (value: DocInfo[K]) => setInfo((i) => ({ ...i, [key]: value }))
  const suggested = suggestFileName({ project: info.project, deliverable: info.deliverable, title: info.title, version: info.version })
  const issues = fileName ? fileNameIssues(fileName) : []

  const copy = async (text: string): Promise<void> => {
    try {
      await navigator.clipboard.writeText(text)
      notify('Copied.')
    } catch {
      notify('The clipboard is not available here. Select the name and copy it.', 'error')
    }
  }

  return (
    <>
      {styles === false && !info.clientTemplate && (
        <Notice kind="warn">
          This document is not built on the BwB template. Adding the BwB styles brings in BwB lists, captions, panels and table
          styles; restyling makes Word&rsquo;s own Normal and heading styles match the BwB format.
          <div className="row" style={{ marginTop: 8 }}>
            <Button small variant="navy" busy={busy === 'styles'}
              onClick={() => void run('styles', async () => {
                await addBwbStyles()
                const changed = supports('1.5') ? await restyleBuiltInStyles() : []
                await refresh()
                return changed.length ? `BwB styles added and ${changed.length} Word styles restyled.` : 'BwB styles added.'
              })}>
              Add BwB styles
            </Button>
          </div>
        </Notice>
      )}
      {styles === true && <Notice>Built on the BwB template: BwB styles are in place.</Notice>}

      <SelectField label="Document type" value={info.docType} onChange={set('docType')}
        options={[{ value: '', label: 'Not set' }, ...DOC_TYPES.map((t) => ({ value: t.id, label: t.name }))]}
        hint="Sets the type-specific checks on the Check tab." />
      <TextField label="Title" value={info.title} onChange={set('title')} />
      <div className="cols-2">
        <TextField label="Project" value={info.project} onChange={set('project')} />
        <TextField label="Deliverable" value={info.deliverable} onChange={set('deliverable')} />
      </div>
      <div className="cols-2">
        <TextField label="Version" value={info.version} onChange={set('version')} placeholder="03" />
        <SelectField<FunderFrame | ''> label="Funder frame" value={info.funder} onChange={set('funder')}
          options={[{ value: '', label: 'Not set' }, ...FUNDER_FRAMES.map((f) => ({ value: f.id, label: f.label }))]} />
      </div>
      <Checkbox label="The client’s template applies" checked={info.clientTemplate} onChange={set('clientTemplate')}
        hint="The client template overrides BwB house style (for example a UNDP template)." />
      <div className="row end">
        <Button variant="primary" disabled={!loaded} busy={busy === 'save'}
          onClick={() => void run('save', async () => {
            await writeDocInfo({ ...info, author: info.author || 'Bankers without Boundaries', company: info.company || 'Bankers without Boundaries' })
            return 'Document details saved.'
          })}>
          Save details
        </Button>
      </div>

      <h3>File name</h3>
      {fileName ? <p className="mono" style={{ margin: '0 0 6px' }}>{fileName}</p> : <p className="muted small">Not saved yet.</p>}
      {issues.map((i) => <Notice key={i.code} kind="warn">{i.message}</Notice>)}
      <div className="field">
        <span className="field-label">Suggested name</span>
        <div className="row">
          <span className="mono grow" style={{ wordBreak: 'break-all' }}>{suggested}</span>
          <Button small onClick={() => void copy(suggested)} aria-label="Copy the suggested file name"><Copy size={13} aria-hidden /> Copy</Button>
        </div>
        <span className="field-hint">YYYYMMDD_Project_Deliverable_Short title_vNN. State lives in the QA hub, so names never need FINAL.</span>
      </div>
    </>
  )
}
