import { useMemo, useState } from 'react'
import { ArrowLeft } from 'lucide-react'
import type { DocType } from '../../brand/docTypes.ts'
import { DOC_TYPES, searchDocTypes } from '../../brand/docTypes.ts'
import { bwbDate, createNewDocument } from '../../word/newDocument.ts'
import { FUNDER_FRAMES } from '../../word/properties.ts'
import type { FunderFrame } from '../../word/properties.ts'
import type { ImageEntry } from '../../brand/images.ts'
import { Button, Checkbox, Disclosure, Notice, SelectField, TextField } from '../components.tsx'
import { ImagePicker } from '../library/ImagePicker.tsx'
import { DocTypeIcon } from '../docTypeIcons.tsx'
import { useAction } from '../useAction.ts'

function TypePicker({ onPick }: { onPick: (t: DocType) => void }) {
  const [query, setQuery] = useState('')
  const types = useMemo(() => searchDocTypes(query), [query])
  const families = [...new Set(types.map((t) => t.family))]
  return (
    <>
      <TextField label="Find a document type" type="search" value={query} onChange={setQuery} placeholder="Inception report, EOI, briefing note…" />
      {families.map((family) => (
        <div key={family}>
          <h3>{family}</h3>
          <div className="tiles">
            {types.filter((t) => t.family === family).map((t) => (
              <button key={t.id} type="button" className="tile" onClick={() => onPick(t)}>
                <span className="tile-mark light" aria-hidden><DocTypeIcon id={t.id} /></span>
                <span className="tile-body">
                  <span className="tile-label">{t.name}</span>
                  <span className="tile-hint" style={{ display: 'block' }}>{t.summary}</span>
                </span>
              </button>
            ))}
          </div>
        </div>
      ))}
      {!types.length && <p className="muted">No match. Try a broader word, or start from a blank report.</p>}
    </>
  )
}

function DetailsForm({ type, onBack }: { type: DocType; onBack: () => void }) {
  const { busy, run } = useAction()
  const [title, setTitle] = useState('')
  const [tagline, setTagline] = useState('')
  const [date, setDate] = useState(bwbDate())
  const [preparedBy, setPreparedBy] = useState('Bankers without Boundaries')
  const [to, setTo] = useState('')
  const [from, setFrom] = useState('')
  const [project, setProject] = useState('')
  const [deliverable, setDeliverable] = useState('')
  const [funder, setFunder] = useState<FunderFrame | ''>('')
  const [disclaimer, setDisclaimer] = useState(type.disclaimer)
  const [dividers, setDividers] = useState(false)
  const [coverPhoto, setCoverPhoto] = useState<ImageEntry | null>(null)
  const [done, setDone] = useState<string[] | null>(null)
  const memo = type.layout === 'memo'

  const create = (): void =>
    void run('create', async () => {
      const details: Array<[string, string]> = memo
        ? ([['To', to], ['From', from], ['Date', date], ['Subject', title]] as Array<[string, string]>).filter(([, v]) => v.trim())
        : []
      const { warnings } = await createNewDocument({
        type, title, tagline, date, preparedBy, details, disclaimer, dividers: dividers && type.layout === 'report', project, deliverable, funder,
        ...(coverPhoto && type.layout === 'report' ? { coverPhoto } : {}),
      })
      setDone(warnings)
      return `${type.name} created in a new window.`
    })

  return (
    <>
      <Button variant="link" small onClick={onBack}><ArrowLeft size={14} aria-hidden /> All document types</Button>
      <h3 style={{ marginTop: 6 }}>{type.name}</h3>
      <p className="section-intro">{type.summary}{type.length ? `. ${type.length}.` : '.'}</p>
      <TextField label="Title" value={title} onChange={setTitle} placeholder={type.name} hint="Title case, no full stop. Two lines at most on the cover." />
      {!memo && <TextField label="Tagline" value={tagline} onChange={setTagline} hint="One sentence under the title on the cover. Optional." />}
      <div className="cols-2">
        <TextField label="Date" value={date} onChange={setDate} hint="D Month YYYY" />
        <TextField label={memo ? 'From' : 'Prepared by'} value={memo ? from : preparedBy} onChange={memo ? setFrom : setPreparedBy} />
      </div>
      {memo && <TextField label="To" value={to} onChange={setTo} />}
      <div className="cols-2">
        <TextField label="Project" value={project} onChange={setProject} placeholder="UNDP Angola" />
        <TextField label="Deliverable" value={deliverable} onChange={setDeliverable} placeholder="D2.19" />
      </div>
      <SelectField label="Funder frame" value={funder} onChange={setFunder}
        options={[{ value: '', label: 'Not set' }, ...FUNDER_FRAMES.map((f) => ({ value: f.id, label: f.label }))]}
        hint="EU grant deliverables get the EU funding acknowledgement and EU checks." />
      {type.layout === 'report' && (
        <>
          <Disclosure title={`Cover photo: ${coverPhoto ? coverPhoto.title : 'the BwB harbour photo'}`}>
            <ImagePicker chooseLabel="Use" onChoose={setCoverPhoto} />
            {coverPhoto && <Button small variant="link" onClick={() => setCoverPhoto(null)}>Use the default photo</Button>}
          </Disclosure>
          <Checkbox label="Open each main section with a navy divider page" checked={dividers} onChange={setDividers}
            hint="For long reports with a few large parts." />
        </>
      )}
      <Checkbox label="End with the BwB disclaimer" checked={disclaimer} onChange={setDisclaimer}
        hint={type.disclaimer ? 'Client-facing documents need it.' : 'Bids and internal notes usually do not.'} />
      <Button variant="primary" block busy={busy === 'create'} onClick={create}>Create document</Button>
      {done && (
        <Notice>
          The new document has opened in its own window. When Word asks to update fields, choose Yes: that fills in the contents
          and the page total. Writing prompts are marked; the Pre-QA check lists any still in place.
          {done.length > 0 && <span className="small" style={{ display: 'block', marginTop: 4 }}>Notes: {done.join(' ')}</span>}
        </Notice>
      )}
    </>
  )
}

export function NewDocument() {
  const [type, setType] = useState<DocType | null>(null)
  return type ? <DetailsForm type={type} onBack={() => setType(null)} /> : <TypePicker onPick={setType} />
}

export const DOC_TYPE_COUNT = DOC_TYPES.length
