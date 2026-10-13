import { useState } from 'react'
import { applyFormat, applyNamedStyle, pickUpFormat } from '../../../ppt/tools/painter.ts'
import type { FormatPart, NamedStyle, ShapeFormat } from '../../../ppt/tools/painterStore.ts'
import {
  FORMAT_PARTS, MAX_STYLES, addStyle, deleteStyle, describeFormat, loadFormat, loadParts, loadStyles, renameStyle, saveParts,
} from '../../../ppt/tools/painterStore.ts'
import { hasPowerPoint, supportsPpt, updateMessage } from '../../../ppt/run.ts'
import { UserFacingError } from '../../../word/run.ts'
import { Button, Checkbox, Section, TextField } from '../../components.tsx'
import { useAction } from '../../useAction.ts'
import './tools.css'

/** The format painter: pick up one shape's look, choose the parts to copy, apply to any slide, keep named styles. */
export function FormatPainterSection() {
  const { busy, run } = useAction()
  const [format, setFormat] = useState<ShapeFormat | null>(() => loadFormat())
  const [parts, setParts] = useState<FormatPart[]>(() => loadParts())
  const [styles, setStyles] = useState<NamedStyle[]>(() => loadStyles())
  const [name, setName] = useState('')
  const [renaming, setRenaming] = useState<{ id: string; name: string } | null>(null)
  const live = hasPowerPoint()
  const reason = (version: string): string | null => (live && !supportsPpt(version) ? updateMessage('This part') : null)
  const blocked = live && !supportsPpt('1.5')
  const off = busy !== null || blocked

  const toggle = (part: FormatPart, on: boolean): void => {
    const next = on ? [...parts, part] : parts.filter((p) => p !== part)
    setParts(next)
    saveParts(next)
  }
  const pickUp = (): void => void run('pick', async () => {
    const message = await pickUpFormat()
    setFormat(loadFormat())
    return message
  })
  const save = (): void => void run('save', async () => {
    if (!format) throw new UserFacingError('Pick up a format first, then save it as a style.')
    setStyles(addStyle(name, format))
    setName('')
    return `Style “${name.trim()}” saved.`
  })
  const rename = (): void => void run('rename', async () => {
    if (!renaming) return
    setStyles(renameStyle(renaming.id, renaming.name))
    setRenaming(null)
    return 'Style renamed.'
  })

  return (
    <Section id="painter" title="Format painter" intro="Pick up the look of one shape, then apply it to shapes on any slide.">
      {blocked && <p className="tools-note">{updateMessage('The format painter')}</p>}
      <div className="tool-row">
        <Button variant="navy" disabled={off} busy={busy === 'pick'} onClick={pickUp}>Pick up format</Button>
        <Button disabled={off || !format} busy={busy === 'apply'} onClick={() => void run('apply', () => applyFormat(parts))}>Apply to selected shapes</Button>
      </div>
      {format
        ? <p className="painter-current">From <b>{`“${format.source}”`}</b>: {describeFormat(format)}.</p>
        : <p className="muted small">Nothing picked up yet.</p>}
      <fieldset className="painter-parts">
        <legend>What to apply</legend>
        {FORMAT_PARTS.map((p) => {
          const why = reason(p.version)
          return <Checkbox key={p.id} label={p.label} checked={parts.includes(p.id)} onChange={(on) => toggle(p.id, on)} {...(why ? { hint: why } : {})} />
        })}
      </fieldset>

      <h3>Saved styles</h3>
      {styles.length > 0 && (
        <ul className="style-list">
          {styles.map((s) => (
            <li key={s.id}>
              {renaming?.id === s.id ? (
                <>
                  <TextField label="New name" value={renaming.name} onChange={(v) => setRenaming({ id: s.id, name: v })} />
                  <Button small variant="navy" disabled={off} onClick={rename}>Save</Button>
                  <Button small variant="link" onClick={() => setRenaming(null)}>Cancel</Button>
                </>
              ) : (
                <>
                  <span className="style-name" title={describeFormat(s.format)}>{s.name}</span>
                  <Button small disabled={off} onClick={() => void run(`style-${s.id}`, () => applyNamedStyle(s.id, parts))}>Apply</Button>
                  <Button small variant="link" onClick={() => setRenaming({ id: s.id, name: s.name })}>Rename</Button>
                  <Button small variant="link" onClick={() => setStyles(deleteStyle(s.id))}>Delete</Button>
                </>
              )}
            </li>
          ))}
        </ul>
      )}
      <TextField label={`Save the picked-up format as a style (up to ${MAX_STYLES})`} value={name} onChange={setName} placeholder="For example: Navy panel" />
      <Button small disabled={busy !== null || !format || styles.length >= MAX_STYLES} onClick={save}>Save as a style</Button>
    </Section>
  )
}
