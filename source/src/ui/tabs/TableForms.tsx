import { useState } from 'react'
import { TABLE_STYLES } from '../../brand/styles.ts'
import type { Block } from '../../word/docxApi.ts'
import { insertBlocks, insertFootnote } from '../../word/insert.ts'
import { supports } from '../../word/run.ts'
import type { TableKind } from '../../word/tables.ts'
import { formatTableAtCursor } from '../../word/tables.ts'
import { Button, Checkbox, SelectField, TextField } from '../components.tsx'
import { useAction } from '../useAction.ts'

const KIND_OPTIONS = (Object.keys(TABLE_STYLES) as TableKind[]).map((k) => ({ value: k, label: TABLE_STYLES[k].label }))

const clamp = (value: string, min: number, max: number, fallback: number): number => {
  const n = Math.round(Number(value))
  return Number.isFinite(n) ? Math.min(max, Math.max(min, n)) : fallback
}

const cells = (text: string): string[] => text.split(/[|\t]/).map((c) => c.trim()).filter(Boolean)

function insertMessage(done: string, warnings: string[]): string {
  return warnings.length ? `${done} ${warnings.join(' ')}` : done
}

export function DataTableForm() {
  const { busy, run } = useAction()
  const [kind, setKind] = useState<TableKind>('standard')
  const [columns, setColumns] = useState('4')
  const [rows, setRows] = useState('4')
  const [header, setHeader] = useState('')
  const [total, setTotal] = useState(false)
  const [caption, setCaption] = useState('')
  const [source, setSource] = useState('')

  const insert = (): void =>
    void run('table', async () => {
      const labels = cells(header)
      const width = labels.length || clamp(columns, 2, 8, 4)
      const height = clamp(rows, 1, 30, 4)
      const body = Array.from({ length: height }, (_, r) =>
        Array.from({ length: width }, (_, c): string | null => (total && r === height - 1 && c === 0 ? 'Total' : null)),
      )
      const block: Block = {
        type: 'table',
        style: kind,
        header: labels.length ? labels : Array.from({ length: width }, (_, c) => `{field:Column ${c + 1}}`),
        rows: body,
        total_row: total,
        caption: caption.trim() || '{field:Table title}',
        ...(source.trim() ? { source: source.trim() } : {}),
      }
      const { warnings } = await insertBlocks([block])
      return insertMessage('Table inserted with its caption.', warnings)
    })

  return (
    <>
      <SelectField label="Kind" value={kind} onChange={setKind} options={KIND_OPTIONS} />
      <div className="cols-2">
        <TextField label="Columns" type="number" value={columns} onChange={setColumns} />
        <TextField label="Body rows" type="number" value={rows} onChange={setRows} />
      </div>
      <TextField label="Column headings" value={header} onChange={setHeader} placeholder="Country | Instrument | Size (USD m) | Year"
        hint="Optional. Separate with a vertical bar; sets the number of columns." />
      <Checkbox label="Last row is a total" checked={total} onChange={setTotal} />
      <TextField label="Caption" value={caption} onChange={setCaption} placeholder="Green bond issuance by country, 2020 to 2025"
        hint="Numbered automatically: Table 1, Table 2." />
      <TextField label="Source" value={source} onChange={setSource} placeholder="BwB analysis of CBI data (2026)" />
      <Button variant="primary" block busy={busy === 'table'} onClick={insert}>Insert table</Button>
    </>
  )
}

export function FormatTable() {
  const { busy, run } = useAction()
  const [kind, setKind] = useState<TableKind>('standard')
  return (
    <>
      <SelectField label="Style" value={kind} onChange={setKind} options={KIND_OPTIONS}
        hint="Pasting from Excel? Paste with Keep Text Only or Use Destination Styles first, so Excel colours do not stay." />
      <Button variant="navy" block busy={busy === 'format'} onClick={() => void run('format', () => formatTableAtCursor(kind))}>
        Format the table at the cursor
      </Button>
    </>
  )
}

export function CaptionForm() {
  const { busy, run } = useAction()
  const [kind, setKind] = useState<'Table' | 'Figure'>('Figure')
  const [text, setText] = useState('')
  const [source, setSource] = useState('')
  return (
    <>
      <SelectField label="Caption for" value={kind} onChange={setKind}
        options={[{ value: 'Figure', label: 'A figure, chart or picture' }, { value: 'Table', label: 'A table' }]}
        hint="Captions go below the table or figure." />
      <TextField label="Title" value={text} onChange={setText} placeholder="Climate finance flows by region, 2025" />
      <TextField label="Source" value={source} onChange={setSource} placeholder="OECD (2026)" />
      <Button variant="primary" block busy={busy === 'caption'}
        onClick={() => void run('caption', async () => {
          const { warnings } = await insertBlocks([{ type: 'caption', kind, text: text.trim() || '{field:Title}', ...(source.trim() ? { source: source.trim() } : {}) }])
          return insertMessage(`${kind} caption inserted.`, warnings)
        })}>
        Insert caption
      </Button>
    </>
  )
}

export function SourceAndFootnote() {
  const { busy, run } = useAction()
  const [source, setSource] = useState('')
  const [note, setNote] = useState('')
  const footnotes = supports('1.5')
  return (
    <>
      <TextField label="Source line" value={source} onChange={setSource} placeholder="Source: IMF World Economic Outlook (2026)"
        hint="8 pt italic charcoal, under a figure or a box." />
      <Button block busy={busy === 'source'}
        onClick={() => void run('source', async () => {
          const text = source.trim() || '{field:Source}'
          const { warnings } = await insertBlocks([{ type: 'source', text: /^source/i.test(text) || text.startsWith('{') ? text : `Source: ${text}` }])
          return insertMessage('Source line inserted.', warnings)
        })}>
        Insert source line
      </Button>
      <h3>Footnote</h3>
      <TextField label="Footnote text" value={note} onChange={setNote} multiline />
      <Button block disabled={!footnotes || !note.trim()} busy={busy === 'footnote'}
        title={footnotes ? undefined : 'Needs a newer Word (WordApi 1.5). Use References, Insert Footnote.'}
        onClick={() => void run('footnote', async () => {
          await insertFootnote(note.trim())
          setNote('')
          return 'Footnote added at the cursor.'
        })}>
        Add footnote at the cursor
      </Button>
    </>
  )
}
