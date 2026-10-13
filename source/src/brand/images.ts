// The BwB image library: photos from the BwB website, converted for Word and PowerPoint by
// scripts/make-image-library.py. The index is loaded only when the picker opens.

export interface ImageEntry {
  id: string
  title: string
  category: string
  tags: string[]
  /** Path of the full-size JPEG, relative to the add-in's root. */
  file: string
  thumb: string
  width: number
  height: number
  bytes: number
}

export interface ImageIndex {
  note: string
  categories: Array<{ id: string; label: string }>
  images: ImageEntry[]
}

let cached: Promise<ImageIndex> | undefined

export function loadImageIndex(url = './assets/images/index.json'): Promise<ImageIndex> {
  cached ??= fetch(url).then((r) => {
    if (!r.ok) throw new Error(`The image library could not be loaded (${r.status}).`)
    return r.json() as Promise<ImageIndex>
  })
  return cached
}

/** Ranks photos for a query by title and tag matches; with no query, keeps the catalogue order. */
export function searchImages(index: ImageIndex, query: string, category = ''): ImageEntry[] {
  const pool = category ? index.images.filter((i) => i.category === category) : index.images
  const words = query.toLowerCase().split(/[\s,]+/).filter(Boolean)
  if (!words.length) return pool
  const scored = pool.map((image) => {
    const title = image.title.toLowerCase()
    let score = 0
    for (const w of words) {
      if (image.tags.includes(w)) score += 3
      else if (image.tags.some((t) => t.startsWith(w) || t.includes(` ${w}`))) score += 2
      if (title.includes(w)) score += 2
    }
    return { image, score }
  })
  return scored.filter((s) => s.score > 0).sort((a, b) => b.score - a.score).map((s) => s.image)
}

export function orientation(image: Pick<ImageEntry, 'width' | 'height'>): 'landscape' | 'portrait' | 'square' {
  const ratio = image.width / image.height
  return ratio > 1.15 ? 'landscape' : ratio < 0.87 ? 'portrait' : 'square'
}

export async function fetchImageBytes(image: ImageEntry): Promise<Uint8Array> {
  const response = await fetch(`./${image.file}`)
  if (!response.ok) throw new Error(`The photo could not be loaded (${response.status}).`)
  return new Uint8Array(await response.arrayBuffer())
}
