import { useEffect, useMemo, useState } from 'react'
import type { ImageEntry, ImageIndex } from '../../brand/images.ts'
import { fetchImageBytes, loadImageIndex, searchImages } from '../../brand/images.ts'
import { currentHost } from '../../host.ts'
import { SLIDE, cmToPt, insertPictureHere } from '../../media/insert.ts'
import { errorMessage } from '../../word/run.ts'
import { Button, Notice, SelectField, TextField } from '../components.tsx'
import { useAction } from '../useAction.ts'

type Size = 'full' | 'half' | 'third'

/** Widths in points: the A4 text width (16.6 cm) in Word, the slide width in PowerPoint. */
function widthFor(size: Size): number {
  const full = currentHost() === 'powerpoint' ? SLIDE.width : cmToPt(16.6)
  return size === 'full' ? full : size === 'half' ? full / 2 : full / 3
}

const SIZES: Array<{ value: Size; label: string }> = [
  { value: 'full', label: 'Full width' },
  { value: 'half', label: 'Half width' },
  { value: 'third', label: 'A third' },
]

export function ImagePicker({ onChoose, chooseLabel }: { onChoose?: (image: ImageEntry) => void; chooseLabel?: string }) {
  const { busy, run } = useAction()
  const [index, setIndex] = useState<ImageIndex | null>(null)
  const [failed, setFailed] = useState<string | null>(null)
  const [query, setQuery] = useState('')
  const [category, setCategory] = useState('')
  const [selected, setSelected] = useState<ImageEntry | null>(null)
  const [size, setSize] = useState<Size>('full')

  useEffect(() => {
    let live = true
    loadImageIndex().then((i) => live && setIndex(i), (e: unknown) => live && setFailed(errorMessage(e)))
    return () => {
      live = false
    }
  }, [])

  const results = useMemo(() => (index ? searchImages(index, query, category) : []), [index, query, category])

  const insert = (image: ImageEntry): void => {
    if (onChoose) {
      onChoose(image)
      return
    }
    void run('image', async () => {
      const widthPt = widthFor(size)
      await insertPictureHere({
        bytes: await fetchImageBytes(image), widthPt, heightPt: (widthPt * image.height) / image.width, altText: image.title,
      })
      return `“${image.title}” inserted.`
    })
  }

  if (failed) return <Notice kind="error">{failed}</Notice>
  if (!index) return <p className="muted small">Loading the image library…</p>
  return (
    <>
      <TextField label={`Search ${index.images.length} BwB photos`} type="search" value={query} onChange={setQuery} placeholder="water, team, energy, city" />
      <div className="row chips" aria-label="Categories">
        {[{ id: '', label: 'All' }, ...index.categories].map((c) => (
          <button key={c.id || 'all'} type="button" className="chip" aria-pressed={category === c.id} onClick={() => setCategory(c.id)}>{c.label}</button>
        ))}
      </div>
      {results.length === 0 && <p className="muted small">No photo matches “{query}”.</p>}
      <div className="image-grid" role="listbox" aria-label="Photos">
        {results.map((image) => (
          <button key={image.id} type="button" role="option" className="image-cell" aria-selected={selected?.id === image.id}
            title={image.title} onClick={() => setSelected(image)} onDoubleClick={() => insert(image)}>
            <img src={`./${image.thumb}`} alt="" loading="lazy" />
            <span>{image.title}</span>
          </button>
        ))}
      </div>
      {!onChoose && <SelectField label="Size" value={size} onChange={setSize} options={SIZES} />}
      <Button variant="primary" block disabled={!selected} busy={busy === 'image'} onClick={() => selected && insert(selected)}>
        {selected ? `${chooseLabel ?? 'Insert'} “${selected.title}”` : 'Choose a photo'}
      </Button>
      <p className="field-hint" style={{ marginTop: 6 }}>Photos from bwb.earth. Team photos are BwB’s own; check a stock photo’s licence before using it outside BwB work.</p>
    </>
  )
}
