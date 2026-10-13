import { useCallback, useEffect, useState } from 'react'
import { hasWord } from '../../word/run.ts'
import type { ReviewState } from '../../word/tools/review.ts'
import {
  DEFAULT_NOTE, acceptAllChanges, addReviewNote, deleteAllComments, readReviewState, rejectAllChanges, toggleTracking,
} from '../../word/tools/review.ts'
import { unavailable } from '../../word/tools/support.ts'
import { Button, Disclosure, Notice, Section, TextField } from '../components.tsx'
import { useAction } from '../useAction.ts'
import { ToolButton } from './ToolButton.tsx'

const NOTES = ['Please check the figures.', 'Source needed.', 'Client to confirm.', 'Check against the terms of reference.']

type Confirm = 'reject' | 'comments' | null

const plural = (n: number, one: string): string => `${n} ${one}${n === 1 ? '' : 's'}`

export function ReviewTools() {
  const { busy, run } = useAction()
  const working = busy !== null
  const [state, setState] = useState<ReviewState | null>(null)
  const [confirm, setConfirm] = useState<Confirm>(null)
  const [note, setNote] = useState('')

  const refresh = useCallback(async (): Promise<void> => {
    setState(await readReviewState().catch(() => null))
  }, [])

  useEffect(() => {
    if (hasWord()) void refresh()
  }, [refresh])

  const act = (key: string, action: () => Promise<string>): void => void run(key, async () => {
    setConfirm(null)
    const message = await action()
    await refresh()
    return message
  })

  const tracking = unavailable('tracking')
  const changes = unavailable('trackedChanges')
  const comments = unavailable('comments')
  const on = state?.tracking && state.tracking !== 'Off'

  return (
    <Section id="review" title="Review" intro="Before sending a document to QA: settle the tracked changes and clear the comments.">
      <div className="review-state">
        <span>Track changes: <strong>{state?.tracking ? (on ? 'on' : 'off') : 'not known'}</strong>.</span>
        {state?.changes !== null && state?.changes !== undefined && <span>{plural(state.changes, 'tracked change')}.</span>}
        {state?.comments !== null && state?.comments !== undefined && <span>{plural(state.comments, 'comment')}.</span>}
        <Button variant="link" small onClick={() => void run('review-refresh', refresh)} disabled={working}>Refresh</Button>
      </div>
      <div className="tool-row">
        <ToolButton label={on ? 'Turn track changes off' : 'Turn track changes on'} variant="navy" busy={busy === 'track'} working={working}
          reason={tracking} onClick={() => act('track', toggleTracking)} />
      </div>
      <div className="tool-row">
        <ToolButton label="Accept all changes" busy={busy === 'accept'} working={working} reason={changes} onClick={() => act('accept', acceptAllChanges)} />
        <ToolButton label="Reject all changes" busy={busy === 'reject'} working={working} reason={changes} onClick={() => setConfirm('reject')} />
        <ToolButton label="Delete all comments" busy={busy === 'comments'} working={working} reason={comments} onClick={() => setConfirm('comments')} />
      </div>
      {confirm && (
        <Notice kind="warn">
          <p className="small">
            {confirm === 'reject'
              ? 'Reject every tracked change in the document? The edits are undone. Ctrl+Z brings them back.'
              : 'Delete every comment in the document, with its replies? Ctrl+Z brings them back.'}
          </p>
          <div className="row">
            <Button small variant="navy" onClick={() => act(confirm, confirm === 'reject' ? rejectAllChanges : deleteAllComments)}>
              {confirm === 'reject' ? 'Reject all' : 'Delete all'}
            </Button>
            <Button small onClick={() => setConfirm(null)}>Cancel</Button>
          </div>
        </Notice>
      )}
      <Disclosure title="Review note">
        <div className="row chips" role="group" aria-label="Quick notes">
          {NOTES.map((n) => (
            <button key={n} type="button" className="chip" aria-pressed={note === n} onClick={() => setNote(n)}>{n}</button>
          ))}
        </div>
        <TextField label="Note for the reviewer" value={note} onChange={setNote} multiline placeholder={DEFAULT_NOTE}
          hint="Added as a comment on the selected text, or at the cursor." />
        <ToolButton label="Add the note" variant="primary" block busy={busy === 'note'} working={working} reason={comments}
          onClick={() => act('note', async () => {
            const message = await addReviewNote(note)
            setNote('')
            return message
          })} />
      </Disclosure>
    </Section>
  )
}
