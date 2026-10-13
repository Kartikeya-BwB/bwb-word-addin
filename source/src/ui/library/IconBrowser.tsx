// The icon browser for the Insert tab in Word and PowerPoint: Lucide (the BwB default), the BwB sectors, Tabler,
// flags and maps, each loaded only when opened or searched, in one fast grid with BwB colours and sizes.

import { useDeferredValue, useEffect, useId, useLayoutEffect, useMemo, useRef, useState } from 'react'
import type { KeyboardEvent, ReactNode } from 'react'
import type { EntryKind, FlagShape, IconEntry, Library, LibraryColour, LibraryId } from '../../brand/iconLibraries.ts'
import {
  LIBRARIES, altTextFor, boxFor, coloursFor, copyTextFor, flagUrl, iconSvgFor, loadLibrary, searchLibraries, sizesFor, strokeSvg,
} from '../../brand/iconLibraries.ts'
import type { Host } from '../../host.ts'
import { currentHost } from '../../host.ts'
import { insertSvgHere } from '../../media/insert.ts'
import { UserFacingError, errorMessage } from '../../word/run.ts'
import { Button, Notice, SelectField, TextField } from '../components.tsx'
import { useAction } from '../useAction.ts'
import './IconBrowser.css'

type Tab = LibraryId | 'all'

const TABS: Array<{ id: Tab; label: string }> = [...LIBRARIES.map(({ id, label }) => ({ id, label })), { id: 'all', label: 'All' }]

// Topics BwB writes about most, from the skill's icon topics; each opens a ready-made search.
const TOPICS = ['climate finance', 'water', 'energy', 'agriculture', 'nature', 'adaptation', 'insurance', 'blended finance', 'gender', 'health', 'governance', 'risk']

// Flag and map regions, in the order the chips show them.
const REGION_ORDER = [
  'Regions', 'Sub-Saharan Africa', 'Middle East and North Africa', 'South Asia', 'East Asia and Pacific', 'Europe and Central Asia',
  'Latin America and the Caribbean', 'North America', 'Antarctica', 'International organisations',
]

const COLOUR_LABELS: Record<LibraryColour, string> = {
  green: 'Green', navy: 'Navy', charcoal: 'Charcoal (notes)', lightGrey: 'Light grey (backdrop)', white: 'White', lime: 'Lime (one accent)',
}

const SHAPES: Array<{ value: FlagShape; label: string }> = [
  { value: 'rect', label: 'Rectangle, 4:3' }, { value: 'square', label: 'Square' }, { value: 'round', label: 'Round' },
]

const NOUN: Record<EntryKind, string> = { stroke: 'Icon', flag: 'Flag', map: 'Map' }

const HINTS: Record<EntryKind, string> = {
  stroke: 'Icons name a topic, mark a callout or label a tile; they are never decoration. Lime and white are for navy panels only.',
  flag: 'Flags keep their own colours.',
  map: 'Maps take the brand colours. Light grey suits a backdrop behind text or markers.',
}

const PLACEHOLDERS: Record<Tab, string> = {
  lucide: 'water, drought, bond, solar', sectors: 'energy, ocean, food', tabler: 'water, drought, bond, solar',
  flags: 'Kenya, UK, EU', maps: 'Kenya, Africa, world', all: 'water, Kenya, solar',
}

const ALL_LICENCES = 'Lucide icons ISC, Tabler icons MIT, flags MIT, maps from Natural Earth via world-atlas (ISC).'

const NOUNS: Record<Tab, string> = { lucide: 'Lucide icons', sectors: 'sector icons', tabler: 'Tabler icons', flags: 'flags', maps: 'maps', all: 'icons, flags and maps' }

// Thumbnails are redrawn as the grid scrolls, so each one is built once per colour.
const thumbs = new Map<string, string>()
function strokeThumb(entry: Extract<IconEntry, { kind: 'stroke' }>, colour: LibraryColour, px: number): string {
  const key = `${entry.key}|${colour}|${px}`
  let uri = thumbs.get(key)
  if (!uri) {
    if (thumbs.size > 5000) thumbs.clear()
    uri = `data:image/svg+xml;charset=utf-8,${encodeURIComponent(strokeSvg(entry.nodes, colour, px))}`
    thumbs.set(key, uri)
  }
  return uri
}

/** Pixels for a crisp picture: four times 96 dpi. */
const pixelsFor = (pt: number): number => Math.max(48, Math.round((pt / 72) * 96 * 4))

async function copyToClipboard(text: string): Promise<void> {
  try {
    await navigator.clipboard.writeText(text)
    return
  } catch {
    // Some Office web views block the clipboard API; a selected text area still copies.
  }
  const area = document.createElement('textarea')
  area.value = text
  area.readOnly = true
  area.className = 'ib-offscreen'
  document.body.append(area)
  area.select()
  const copied = document.execCommand('copy')
  area.remove()
  if (!copied) throw new UserFacingError(`The name could not be copied. It is “${text}”.`)
}

function Thumb({ entry, colour, shape, px }: { entry: IconEntry; colour: LibraryColour; shape: FlagShape; px: number }) {
  if (entry.kind === 'stroke') return <img src={strokeThumb(entry, colour, px)} alt="" width={px} height={px} />
  if (entry.kind === 'flag') {
    return <img src={flagUrl(entry.id, shape)} alt="" loading="lazy" className={`ib-flag ${shape}`} style={{ width: shape === 'rect' ? px * 1.4 : px * 1.05 }} />
  }
  const mask = `url("./${entry.file}")`
  return <span className={`ib-map ib-c-${colour}`} style={{ width: px * 1.4, height: px * 1.4, maskImage: mask, WebkitMaskImage: mask }} />
}

const GAP = 4
const MIN_CELL = 46
const VIEW = 288

/** A grid that draws only the rows in view, so that 5,000 icons scroll as fast as 50. Arrow keys move the choice. */
function IconGrid({ entries, selected, onSelect, onActivate, cell: renderCell, onNavy, label }: {
  entries: IconEntry[]
  selected: IconEntry | null
  onSelect: (entry: IconEntry) => void
  onActivate: (entry: IconEntry) => void
  cell: (entry: IconEntry) => ReactNode
  onNavy: boolean
  label: string
}) {
  const ref = useRef<HTMLDivElement>(null)
  const id = useId()
  const [width, setWidth] = useState(0)
  const [top, setTop] = useState(0)

  useLayoutEffect(() => {
    const el = ref.current
    if (!el) return
    setWidth(el.clientWidth)
    if (typeof ResizeObserver === 'undefined') return
    const observer = new ResizeObserver(() => setWidth(el.clientWidth))
    observer.observe(el)
    return () => observer.disconnect()
  }, [])

  useEffect(() => {
    if (ref.current) ref.current.scrollTop = 0
    setTop(0)
  }, [entries])

  const cols = Math.max(1, Math.floor((width + GAP) / (MIN_CELL + GAP)))
  const size = width ? (width - GAP * (cols - 1)) / cols : MIN_CELL
  const rowHeight = size + GAP
  const rows = Math.ceil(entries.length / cols)
  const firstRow = Math.max(0, Math.floor(top / rowHeight) - 2)
  const lastRow = Math.min(rows, Math.ceil((top + VIEW) / rowHeight) + 2)
  const index = selected ? entries.findIndex((e) => e.key === selected.key) : -1
  const shown = index >= firstRow * cols && index < lastRow * cols

  const move = (next: number): void => {
    const at = Math.max(0, Math.min(entries.length - 1, next))
    const target = entries[at]
    const el = ref.current
    if (!target || !el) return
    onSelect(target)
    const row = Math.floor(at / cols)
    if (row * rowHeight < el.scrollTop) el.scrollTop = row * rowHeight
    else if ((row + 1) * rowHeight > el.scrollTop + el.clientHeight) el.scrollTop = (row + 1) * rowHeight - el.clientHeight
  }

  const onKey = (e: KeyboardEvent<HTMLDivElement>): void => {
    const page = cols * Math.max(1, Math.floor(VIEW / rowHeight))
    const steps: Partial<Record<string, number>> = { ArrowRight: 1, ArrowLeft: -1, ArrowDown: cols, ArrowUp: -cols, PageDown: page, PageUp: -page }
    const step = steps[e.key]
    if (e.key === 'Enter' && selected) onActivate(selected)
    else if (e.key === 'Home') move(0)
    else if (e.key === 'End') move(entries.length - 1)
    else if (step !== undefined) move(index < 0 ? 0 : index + step)
    else return
    e.preventDefault()
  }

  return (
    <div ref={ref} className={`ib-grid ${onNavy ? 'on-navy' : ''}`} role="listbox" tabIndex={0} aria-label={label}
      aria-activedescendant={shown ? `${id}-${index}` : undefined} onScroll={(e) => setTop(e.currentTarget.scrollTop)} onKeyDown={onKey}>
      <div className="ib-grid-space" style={{ height: rows * rowHeight - GAP }}>
        <div className="ib-grid-rows" style={{ top: firstRow * rowHeight, gridTemplateColumns: `repeat(${cols}, 1fr)`, gap: GAP }}>
          {entries.slice(firstRow * cols, lastRow * cols).map((entry, i) => {
            const n = firstRow * cols + i
            return (
              <div key={entry.key} id={`${id}-${n}`} role="option" aria-selected={selected?.key === entry.key} aria-setsize={entries.length}
                aria-posinset={n + 1} aria-label={entry.name} className="ib-cell" title={entry.name} style={{ height: size }}
                onClick={() => onSelect(entry)} onDoubleClick={() => onActivate(entry)}>
                {renderCell(entry)}
              </div>
            )
          })}
        </div>
      </div>
    </div>
  )
}

export function IconBrowser({ host = currentHost() }: { host?: Host }) {
  const { busy, run } = useAction()
  const [tab, setTab] = useState<Tab>('lucide')
  const [libraries, setLibraries] = useState<Partial<Record<LibraryId, Library>>>({})
  const [failed, setFailed] = useState<string | null>(null)
  const [query, setQuery] = useState('')
  const search = useDeferredValue(query).trim()
  const [group, setGroup] = useState('')
  const [selected, setSelected] = useState<IconEntry | null>(null)
  const [onNavy, setOnNavy] = useState(false)
  const [colour, setColour] = useState<LibraryColour>('green')
  const [sizeId, setSizeId] = useState('')
  const [customPt, setCustomPt] = useState('72')
  const [shape, setShape] = useState<FlagShape>('rect')

  // Only what is open or searched is loaded: one library, or every library once "All" has a query.
  const wanted = tab === 'all' ? (search ? LIBRARIES.map((l) => l.id).join(' ') : '') : tab
  useEffect(() => {
    let live = true
    for (const id of wanted.split(' ').filter(Boolean) as LibraryId[]) {
      loadLibrary(id).then(
        (library) => live && setLibraries((loaded) => (loaded[id] ? loaded : { ...loaded, [id]: library })),
        (e: unknown) => live && setFailed(errorMessage(e)),
      )
    }
    return () => {
      live = false
    }
  }, [wanted])

  const ready = wanted.split(' ').filter(Boolean).every((id) => libraries[id as LibraryId])
  const results = useMemo(() => {
    const open = wanted.split(' ').filter(Boolean).flatMap((id) => libraries[id as LibraryId] ?? [])
    const inGroup = (e: IconEntry): boolean => !group || e.group === group
    if (search) return searchLibraries(open, search).filter(inGroup)
    return tab === 'all' ? [] : (open[0]?.entries ?? []).filter(inGroup)
  }, [libraries, wanted, search, group, tab])

  const groups = tab === 'flags' || tab === 'maps'
    ? [...new Set((libraries[tab]?.entries ?? []).map((e) => e.group))].sort((a, b) => REGION_ORDER.indexOf(a) - REGION_ORDER.indexOf(b))
    : []

  const kind: EntryKind = selected?.kind ?? (tab === 'flags' ? 'flag' : tab === 'maps' ? 'map' : 'stroke')
  const colourFor = (k: EntryKind): LibraryColour => {
    const allowed = coloursFor(k, onNavy)
    return allowed.includes(colour) ? colour : allowed[0] ?? 'green'
  }
  const sizes = sizesFor(kind, host)
  const sizeOptions = [...sizes.map((s) => ({ value: s.id, label: s.label })), ...(host === 'powerpoint' ? [{ value: 'custom', label: 'Custom' }] : [])]
  const fallbackSize = host === 'word' && kind === 'stroke' ? 'callout' : 'medium'
  const size = sizeOptions.some((o) => o.value === sizeId) ? sizeId : fallbackSize
  const custom = Math.min(960, Math.max(6, Number(customPt) || 72))
  const pt = size === 'custom' ? custom : sizes.find((s) => s.id === size)?.pt ?? 48

  const choose = (next: Tab): void => {
    setTab(next)
    setGroup('')
    setSelected(null)
    setFailed(null)
  }

  const insert = (entry: IconEntry): void => {
    setSelected(entry)
    if (busy) return
    void run('insert', async () => {
      const box = boxFor(entry, pt, shape)
      const svg = await iconSvgFor(entry, colourFor(entry.kind), pixelsFor(box.widthPt), shape)
      await insertSvgHere({ svg, ...box, altText: altTextFor(entry) })
      return `${NOUN[entry.kind]} “${entry.name}” inserted.`
    })
  }

  const copy = (entry: IconEntry): void =>
    void run('copy', async () => {
      const text = copyTextFor(entry)
      await copyToClipboard(text)
      return `Copied “${text}”.`
    })

  const library = LIBRARIES.find((l) => l.id === tab)
  const total = library ? libraries[library.id]?.entries.length : undefined
  const noun = NOUNS[tab]
  const searchLabel = tab === 'all' ? 'Search every library' : `Search ${total ? `${total.toLocaleString('en-GB')} ` : ''}${noun}`
  const sizeHint = kind === 'map' ? 'The longer side.' : kind === 'flag' ? 'The width.'
    : host === 'word' ? 'In text, a callout header, a tile or a panel.' : 'The width and height.'
  const navyBackground = onNavy && kind !== 'flag'

  return (
    <>
      {failed && <Notice kind="error">{failed}</Notice>}
      <div className="ib-tabs" role="group" aria-label="Icon libraries">
        {TABS.map((t) => (
          <button key={t.id} type="button" className="ib-tab" aria-pressed={tab === t.id} onClick={() => choose(t.id)}>{t.label}</button>
        ))}
      </div>
      <TextField label={searchLabel} type="search" value={query} onChange={setQuery} placeholder={PLACEHOLDERS[tab]} />
      {groups.length > 0 ? (
        <div className="row chips" aria-label="Regions">
          {['', ...groups].map((g) => (
            <button key={g || 'all'} type="button" className="chip" aria-pressed={group === g} onClick={() => setGroup(g)}>{g || 'All regions'}</button>
          ))}
        </div>
      ) : tab !== 'sectors' && tab !== 'flags' && tab !== 'maps' && (
        <div className="row chips" aria-label="Topics">
          {TOPICS.map((t) => (
            <button key={t} type="button" className="chip" aria-pressed={query === t} onClick={() => setQuery(t)}>{t}</button>
          ))}
        </div>
      )}

      {!ready && !failed && <p className="muted small">Loading {tab === 'all' ? 'every library' : noun}…</p>}
      {ready && tab === 'all' && !search && <p className="muted small">Type a word to search icons, flags and maps at once.</p>}
      {ready && search && !results.length && <p className="muted small">Nothing matches “{search}”. Try a broader word.</p>}
      {ready && results.length > 0 && (
        <>
          <p className="muted small ib-count" aria-live="polite">
            {search ? `${results.length.toLocaleString('en-GB')} ${results.length === 1 ? 'match' : 'matches'}` : `${results.length.toLocaleString('en-GB')} ${noun}`}
          </p>
          <IconGrid entries={results} selected={selected} onSelect={setSelected} onActivate={insert} onNavy={navyBackground}
            label={noun}
            cell={(entry) => <Thumb entry={entry} colour={colourFor(entry.kind)} shape={shape} px={24} />} />
        </>
      )}

      {selected && (
        <div className={`ib-detail ${navyBackground ? 'on-navy' : ''}`}>
          <span className="ib-detail-mark"><Thumb entry={selected} colour={colourFor(selected.kind)} shape={shape} px={40} /></span>
          <span className="ib-detail-text">
            <strong>{selected.name}</strong>
            <span>{[LIBRARIES.find((l) => l.id === selected.library)?.label, selected.group].filter(Boolean).join(', ')}</span>
          </span>
        </div>
      )}
      {kind !== 'flag' && (
        <div className="row chips" role="group" aria-label="Background">
          <button type="button" className="chip" aria-pressed={!onNavy} onClick={() => setOnNavy(false)}>On white</button>
          <button type="button" className="chip" aria-pressed={onNavy} onClick={() => setOnNavy(true)}>On a navy panel</button>
        </div>
      )}
      <div className="cols-2">
        {kind === 'flag'
          ? <SelectField label="Shape" value={shape} onChange={setShape} options={SHAPES} />
          : <SelectField label="Colour" value={colourFor(kind)} onChange={setColour} options={coloursFor(kind, onNavy).map((c) => ({ value: c, label: COLOUR_LABELS[c] }))} />}
        <SelectField label="Size" value={size} onChange={setSizeId} options={sizeOptions} hint={sizeHint} />
      </div>
      {size === 'custom' && <TextField label="Custom size in points" type="number" value={customPt} onChange={setCustomPt} hint="From 6 to 960 points." />}
      <Button variant="primary" block disabled={!selected} busy={busy === 'insert'} onClick={() => selected && insert(selected)}>
        {selected ? `Insert “${selected.name}”` : `Choose ${kind === 'flag' ? 'a flag' : kind === 'map' ? 'a map' : 'an icon'}`}
      </Button>
      <div className="row ib-actions">
        <Button variant="link" small disabled={!selected} busy={busy === 'copy'} onClick={() => selected && copy(selected)}>Copy name</Button>
        <span className="muted small">Double-click to insert</span>
      </div>
      <p className="field-hint">{HINTS[kind]} {library?.licence ?? ALL_LICENCES}</p>
    </>
  )
}
