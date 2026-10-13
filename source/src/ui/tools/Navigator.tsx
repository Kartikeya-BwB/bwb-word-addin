import { useCallback, useEffect, useRef, useState } from 'react'
import { Heading, Image, RefreshCw, Table } from 'lucide-react'
import { errorMessage, hasWord } from '../../word/run.ts'
import type { OutlineItem, OutlineKind } from '../../word/tools/navigator.ts'
import { goToItem, insertCrossReference, readCursorHeading, readOutline, shiftHeadingLevel } from '../../word/tools/navigator.ts'
import { unavailable } from '../../word/tools/support.ts'
import { Button, Checkbox, Notice, Section } from '../components.tsx'
import { useAction } from '../useAction.ts'

type Filter = 'all' | OutlineKind

const FILTERS: Array<{ value: Filter; label: string }> = [
  { value: 'all', label: 'All' },
  { value: 'heading', label: 'Headings' },
  { value: 'table', label: 'Tables' },
  { value: 'figure', label: 'Figures' },
]

const ICONS = { heading: Heading, table: Table, figure: Image } as const

/** Calls `onChange` when the selection in Word moves, if this Word reports it. Returns a function that stops. */
function watchSelection(onChange: () => void): () => void {
  try {
    const doc = Office.context.document
    const type = Office.EventType.DocumentSelectionChanged
    if (typeof doc.addHandlerAsync !== 'function' || !type) return () => undefined
    doc.addHandlerAsync(type, onChange)
    return () => {
      try {
        doc.removeHandlerAsync(type, { handler: onChange })
      } catch {
        // The pane is closing.
      }
    }
  } catch {
    return () => undefined
  }
}

export function Navigator() {
  const { busy, run } = useAction()
  const working = busy !== null
  const [outline, setOutline] = useState<OutlineItem[] | null>(null)
  const [failed, setFailed] = useState<string | null>(null)
  const [filter, setFilter] = useState<Filter>('all')
  const [cursor, setCursor] = useState<{ level: number; text: string } | null>(null)
  const [showNumber, setShowNumber] = useState(false)
  const timer = useRef<number | undefined>(undefined)
  const reason = unavailable('core')
  const crossRef = unavailable('crossReference')

  const refresh = useCallback(async (): Promise<void> => {
    try {
      setOutline(await readOutline())
      setFailed(null)
    } catch (error) {
      setFailed(errorMessage(error))
    }
  }, [])

  const readCursor = useCallback((): void => {
    readCursorHeading().then(setCursor, () => setCursor(null))
  }, [])

  useEffect(() => {
    if (!hasWord() || reason) return undefined
    void refresh()
    readCursor()
    const stop = watchSelection(() => {
      window.clearTimeout(timer.current)
      timer.current = window.setTimeout(readCursor, 300)
    })
    return () => {
      window.clearTimeout(timer.current)
      stop()
    }
  }, [refresh, readCursor, reason])

  const shift = (delta: 1 | -1): void => void run(`shift${delta}`, async () => {
    const message = await shiftHeadingLevel(delta)
    readCursor()
    await refresh()
    return message
  })

  const shown = (outline ?? []).filter((i) => filter === 'all' || i.kind === filter)

  return (
    <Section id="navigator" title="Navigator"
      intro="Every heading, table and figure in the document. Choose one to go there."
      action={<Button small onClick={() => void run('refresh', async () => { await refresh() })} busy={busy === 'refresh'} disabled={Boolean(reason)}
        title="Read the outline again"><RefreshCw size={13} aria-hidden /> Refresh</Button>}>
      {reason && <Notice kind="warn">{reason}</Notice>}
      <div className="cursor-status">
        <span className="grow">
          {cursor ? <>At the cursor: <strong>Heading {cursor.level}</strong> {cursor.text}</> : 'The cursor is not in a heading.'}
        </span>
        <Button small disabled={working || !cursor || Boolean(reason)} onClick={() => shift(-1)} title="Make the heading at the cursor one level higher">Promote</Button>
        <Button small disabled={working || !cursor || Boolean(reason)} onClick={() => shift(1)} title="Make the heading at the cursor one level lower">Demote</Button>
      </div>
      <div className="row chips" role="group" aria-label="Show">
        {FILTERS.map((f) => (
          <button key={f.value} type="button" className="chip" aria-pressed={filter === f.value} onClick={() => setFilter(f.value)}>{f.label}</button>
        ))}
      </div>
      {failed && <Notice kind="error">{failed}</Notice>}
      {outline && !shown.length && <p className="muted small">Nothing to show yet: add headings in the BwB styles, and captions to tables and figures.</p>}
      {shown.length > 0 && (
        <ul className="outline">
          {shown.map((item) => {
            const Icon = ICONS[item.kind]
            return (
              <li key={item.id}>
                <button type="button" className={`outline-item ${item.kind} level-${Math.min(item.level, 5)}`} title={item.label}
                  onClick={() => void run(`go-${item.id}`, async () => { await goToItem(item) })}>
                  <Icon size={13} aria-hidden />
                  <span>{item.label}</span>
                </button>
                {!crossRef && (
                  <button type="button" className="outline-ref" disabled={working || item.paragraph < 0}
                    title={item.paragraph < 0 ? 'Add a caption to this table to refer to it' : 'Insert a cross-reference to this at the cursor'}
                    onClick={() => void run(`ref-${item.id}`, () => insertCrossReference(item, showNumber ? 'number' : 'text'))}>
                    Ref
                  </button>
                )}
              </li>
            )
          })}
        </ul>
      )}
      {crossRef ? (
        <p className="field-hint">{crossRef}</p>
      ) : (
        <>
          <p className="field-hint">To insert a cross-reference, put the cursor where it goes, then choose Ref next to the heading, table or figure.</p>
          <Checkbox label="For headings, show the number (2.1) rather than the text" checked={showNumber} onChange={setShowNumber} />
        </>
      )}
    </Section>
  )
}
