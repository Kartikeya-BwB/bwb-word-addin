import { useEffect, useMemo, useState } from 'react'
import type { IconColourName, IconIndex, IconSize } from '../../brand/icons.ts'
import { ICON_SIZES_CM, iconSvg, loadIconIndex, searchIcons } from '../../brand/icons.ts'
import { svgToPng } from '../../word/iconRenderer.ts'
import { insertPicture } from '../../word/marks.ts'
import { errorMessage } from '../../word/run.ts'
import { Button, Notice, SelectField, TextField } from '../components.tsx'
import { useAction } from '../useAction.ts'

// Topics BwB writes about most, from the skill's icon topics; each opens a ready-made search.
const TOPICS = ['climate finance', 'water', 'energy', 'agriculture', 'nature', 'adaptation', 'insurance', 'blended finance', 'gender', 'health', 'governance', 'risk']

const COLOURS: Array<{ value: Extract<IconColourName, 'green' | 'navy' | 'charcoal'>; label: string }> = [
  { value: 'green', label: 'Green' },
  { value: 'navy', label: 'Navy' },
  { value: 'charcoal', label: 'Charcoal (notes)' },
]

const SIZES: Array<{ value: IconSize; label: string }> = [
  { value: 'inline', label: 'In text, 0.4 cm' },
  { value: 'callout', label: 'Callout, 0.6 cm' },
  { value: 'marker', label: 'Marker, 0.8 cm' },
  { value: 'panel', label: 'Panel, 1.2 cm' },
]

const dataUri = (svg: string): string => `data:image/svg+xml;charset=utf-8,${encodeURIComponent(svg)}`

/** Pixels for a crisp picture at the brand size: four times 96 dpi. */
const pixelsFor = (cm: number): number => Math.round((cm / 2.54) * 96 * 4)

export function IconPicker() {
  const { busy, run } = useAction()
  const [index, setIndex] = useState<IconIndex | null>(null)
  const [failed, setFailed] = useState<string | null>(null)
  const [query, setQuery] = useState('')
  const [selected, setSelected] = useState<string | null>(null)
  const [colour, setColour] = useState<(typeof COLOURS)[number]['value']>('green')
  const [size, setSize] = useState<IconSize>('callout')

  useEffect(() => {
    let live = true
    loadIconIndex().then(
      (i) => live && setIndex(i),
      (e: unknown) => live && setFailed(errorMessage(e)),
    )
    return () => {
      live = false
    }
  }, [])

  const hits = useMemo(() => (index && query.trim() ? searchIcons(index, query, 48) : []), [index, query])
  const total = index ? Object.keys(index.icons).length : 0

  const insert = (name: string): void =>
    void run('icon', async () => {
      if (!index) return
      const cm = ICON_SIZES_CM[size]
      const px = pixelsFor(cm)
      const svg = iconSvg(index, name, colour, px)
      if (!svg) throw new Error(`The icon “${name}” could not be drawn.`)
      await insertPicture(await svgToPng(svg, px), cm, cm, `Icon: ${name.replace(/-/g, ' ')}`)
      return `Icon “${name}” inserted at ${cm} cm.`
    })

  if (failed) return <Notice kind="error">{failed}</Notice>
  return (
    <>
      <TextField label={`Search ${total ? total.toLocaleString('en-GB') : ''} Lucide icons`} type="search" value={query} onChange={setQuery}
        placeholder="water, drought, bond, solar" />
      <div className="row chips" aria-label="Topics">
        {TOPICS.map((t) => (
          <button key={t} type="button" className="chip" aria-pressed={query === t} onClick={() => setQuery(t)}>{t}</button>
        ))}
      </div>
      {index && query.trim() && !hits.length && <p className="muted small">No icon matches “{query}”. Try a broader word.</p>}
      {index && hits.length > 0 && (
        <div className="icon-grid" role="listbox" aria-label="Icons">
          {hits.map(({ name }) => {
            const svg = iconSvg(index, name, colour, 22)
            return svg ? (
              <button key={name} type="button" role="option" className="icon-cell" title={name} aria-label={name}
                aria-selected={selected === name} aria-pressed={selected === name}
                onClick={() => setSelected(name)} onDoubleClick={() => insert(name)}>
                <img src={dataUri(svg)} alt="" width={22} height={22} />
              </button>
            ) : null
          })}
        </div>
      )}
      <div className="cols-2" style={{ marginTop: 10 }}>
        <SelectField label="Colour" value={colour} onChange={setColour} options={COLOURS} />
        <SelectField label="Size" value={size} onChange={setSize} options={SIZES}
          hint="In text, a callout header, a tile or section marker, or a large panel icon." />
      </div>
      <Button variant="primary" block disabled={!selected} busy={busy === 'icon'} onClick={() => selected && insert(selected)}>
        {selected ? `Insert “${selected}”` : 'Choose an icon'}
      </Button>
      <p className="field-hint" style={{ marginTop: 6 }}>
        Icons name a topic, mark a callout or label a tile; they are never decoration. Lime and white are for navy panels only, so they are
        not offered here. Lucide icons, ISC licence.
      </p>
    </>
  )
}
