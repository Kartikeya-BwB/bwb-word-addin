import { useState } from 'react'
import type { Block } from '../../word/docxApi.ts'
import { insertBlocks } from '../../word/insert.ts'
import { Button, SelectField, TextField } from '../components.tsx'
import { useAction } from '../useAction.ts'

const lines = (text: string): string[] => text.split('\n').map((l) => l.trim()).filter(Boolean)

function useInsert(): { busy: string | null; insert: (key: string, blocks: Block[], done: string) => void } {
  const { busy, run } = useAction()
  const insert = (key: string, blocks: Block[], done: string): void =>
    void run(key, async () => {
      const { warnings } = await insertBlocks(blocks)
      return warnings.length ? `${done} ${warnings.join(' ')}` : done
    })
  return { busy, insert }
}

export function CalloutForm() {
  const { busy, insert } = useInsert()
  const [variant, setVariant] = useState<'green' | 'charcoal'>('green')
  const [title, setTitle] = useState('')
  const [text, setText] = useState('')
  const [bullets, setBullets] = useState('')
  const [icon, setIcon] = useState('')
  return (
    <>
      <SelectField label="Kind" value={variant} onChange={setVariant}
        options={[{ value: 'green', label: 'Green: a takeaway, tip or lesson' }, { value: 'charcoal', label: 'Charcoal: a caveat or note' }]} />
      <TextField label="Title" value={title} onChange={setTitle} placeholder="Key takeaway" />
      <TextField label="Text" value={text} onChange={setText} multiline />
      <TextField label="Bullets (one per line)" value={bullets} onChange={setBullets} multiline hint="Optional. Green circles inside the callout." />
      <TextField label="Icon" value={icon} onChange={setIcon} placeholder="lightbulb" hint="Optional. A Lucide icon name for the header; search for one under Icons." />
      <Button variant="primary" block busy={busy === 'callout'}
        onClick={() => insert('callout', [{
          type: 'callout', variant, title: title || '{field:Title}', text: text || '{field:Key message}', bullets: lines(bullets),
          ...(icon.trim() ? { icon: icon.trim() } : {}),
        }], 'Callout inserted.')}>
        Insert callout
      </Button>
    </>
  )
}

export function KeyFiguresForm() {
  const { busy, insert } = useInsert()
  const [items, setItems] = useState([{ value: '', label: '' }, { value: '', label: '' }, { value: '', label: '' }])
  const [hero, setHero] = useState('0')
  const update = (i: number, key: 'value' | 'label', v: string): void =>
    setItems((all) => all.map((it, j) => (j === i ? { ...it, [key]: v } : it)))
  return (
    <>
      <p className="section-intro">Two to four headline numbers on a navy panel. Write currencies as a code: USD 250m.</p>
      {items.map((it, i) => (
        <div className="cols-2" key={i}>
          <TextField label={`Figure ${i + 1}`} value={it.value} onChange={(v) => update(i, 'value', v)} placeholder="USD 102bn" />
          <TextField label="Label" value={it.label} onChange={(v) => update(i, 'label', v)} placeholder="Current flow" />
        </div>
      ))}
      <div className="row">
        <Button small disabled={items.length >= 4} onClick={() => setItems((a) => [...a, { value: '', label: '' }])}>Add a figure</Button>
        <Button small disabled={items.length <= 2} onClick={() => setItems((a) => a.slice(0, -1))}>Remove the last</Button>
      </div>
      <SelectField label="Figure shown in lime" value={hero} onChange={setHero}
        options={[{ value: '-1', label: 'None' }, ...items.map((_, i) => ({ value: String(i), label: `Figure ${i + 1}` }))]} />
      <Button variant="primary" block busy={busy === 'stats'}
        onClick={() => insert('stats', [{
          type: 'stats',
          items: items.map((it, i) => ({ value: it.value || `{field:Figure ${i + 1}}`, label: it.label || '{field:Label}' })),
          ...(hero !== '-1' ? { hero: Number(hero) } : {}),
        }], 'Key figures inserted.')}>
        Insert key figures
      </Button>
    </>
  )
}

export function PullQuoteForm() {
  const { busy, insert } = useInsert()
  const [text, setText] = useState('')
  const [attribution, setAttribution] = useState('')
  return (
    <>
      <TextField label="Quotation" value={text} onChange={setText} multiline hint="One short quotation that carries a point. It is left word for word." />
      <TextField label="Attribution" value={attribution} onChange={setAttribution} placeholder="Name, role, organisation" />
      <Button variant="primary" block busy={busy === 'quote'}
        onClick={() => insert('quote', [{ type: 'pull_quote', text: text || '{field:Quotation}', ...(attribution ? { attribution } : {}) }], 'Pull quote inserted.')}>
        Insert pull quote
      </Button>
    </>
  )
}

export function SnapshotForm() {
  const { busy, insert } = useInsert()
  const [rows, setRows] = useState('')
  return (
    <>
      <TextField label="Facts (one per line: eyebrow | value | note)" value={rows} onChange={setRows} multiline
        placeholder={'Country | Angola | Population 37m\nSector | Green bonds | First issue 2026'} hint="A navy grid of key facts for the start of a briefing. Three per row." />
      <Button variant="primary" block busy={busy === 'snapshot'}
        onClick={() => {
          const items = lines(rows).map((l) => {
            const [eyebrow = '', value = '', note = ''] = l.split('|').map((x) => x.trim())
            return { eyebrow, value, note }
          })
          insert('snapshot', [{ type: 'snapshot', items: items.length ? items : [{ eyebrow: '{field:Eyebrow}', value: '{field:Value}', note: '{field:Note}' }] }], 'Snapshot inserted.')
        }}>
        Insert snapshot
      </Button>
    </>
  )
}
