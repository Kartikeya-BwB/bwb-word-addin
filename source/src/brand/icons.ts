// The Lucide icon library: loading the index, resolving older names and ranking search results.
// Search follows the BwB Document Format skill (scripts/bwb_lib/icons.py): names, tags and BwB topics.

export type IconNode = [string, Record<string, string>]

export interface IconIndex {
  version: string
  icons: Record<string, { n: IconNode[]; t: string[] }>
  aliases: Record<string, string>
  topics: Record<string, string[]>
}

export const ICON_HEX = { navy: '022338', green: '328840', lime: 'BED62F', charcoal: '4A4A4C', white: 'FFFFFF' } as const
export type IconColourName = keyof typeof ICON_HEX

/** Brand sizes in centimetres: inline with text, callout header, tile or marker, large panel icon. */
export const ICON_SIZES_CM = { inline: 0.4, callout: 0.6, marker: 0.8, panel: 1.2 } as const
export type IconSize = keyof typeof ICON_SIZES_CM

let cached: Promise<IconIndex> | undefined

export function loadIconIndex(url = './assets/lucide/index.json'): Promise<IconIndex> {
  cached ??= fetch(url).then((r) => {
    if (!r.ok) throw new Error(`The icon library could not be loaded (${r.status}).`)
    return r.json() as Promise<IconIndex>
  })
  return cached
}

export function normaliseIconName(name: string): string {
  return name.trim().toLowerCase().replace(/[_ ]/g, '-').replace(/^lucide[-:]/, '').replace(/-icon$/, '')
}

export function resolveIcon(index: IconIndex, name: string): string | null {
  const n = normaliseIconName(name)
  if (index.icons[n]) return n
  return index.aliases[n] ?? null
}

function escapeAttr(value: string): string {
  return value.replace(/&/g, '&amp;').replace(/"/g, '&quot;').replace(/</g, '&lt;')
}

/** Standalone SVG markup for an icon, stroked in a BwB colour. */
export function iconSvg(index: IconIndex, name: string, colour: IconColourName, px = 24): string | null {
  const resolved = resolveIcon(index, name)
  const icon = resolved ? index.icons[resolved] : undefined
  if (!icon) return null
  const body = icon.n
    .map(([tag, attrs]) => `<${tag} ${Object.entries(attrs).map(([k, v]) => `${k}="${escapeAttr(v)}"`).join(' ')}/>`)
    .join('')
  return (
    `<svg xmlns="http://www.w3.org/2000/svg" width="${px}" height="${px}" viewBox="0 0 24 24" fill="none" ` +
    `stroke="#${ICON_HEX[colour]}" stroke-width="2" stroke-linecap="round" stroke-linejoin="round">${body}</svg>`
  )
}

export interface IconHit {
  name: string
  tags: string[]
  score: number
}

/** Ranks icons for a query: whole and partial name matches, tags, then BwB topics such as "climate finance". */
export function searchIcons(index: IconIndex, query: string, limit = 60): IconHit[] {
  const q = normaliseIconName(query).replace(/-/g, ' ')
  const words = q.split(/[\s,;/]+/).filter(Boolean)
  if (!words.length) return []
  const topicScore = new Map<string, number>()
  for (const [key, names] of Object.entries(index.topics)) {
    const k = key.toLowerCase()
    const parts = k.split(' ')
    const hit = k === q || words.includes(k) || (parts.length > 1 && parts.every((p) => words.includes(p)))
    if (!hit) continue
    names.forEach((n, pos) => {
      const r = resolveIcon(index, n)
      if (r) topicScore.set(r, Math.max(topicScore.get(r) ?? 0, 70 - pos * 0.6))
    })
  }
  const out: IconHit[] = []
  for (const [name, { t: tags }] of Object.entries(index.icons)) {
    const parts = name.split('-')
    let matched = 0
    let score = 0
    for (const w of words) {
      let s = 0
      if (name === w) s = 100
      else if (parts.includes(w)) s = 60
      else if (name.startsWith(w)) s = 45
      else if (w.length > 2 && name.includes(w)) s = 30
      else if (tags.includes(w)) s = 40
      else if (w.length > 3 && tags.some((t) => t.startsWith(w) || t.includes(w))) s = 20
      if (s) {
        matched += 1
        score += s
      }
    }
    const topic = topicScore.get(name)
    if (topic !== undefined) {
      matched = Math.max(matched, 1) + (matched ? 1 : 0)
      score += topic
    }
    if (matched) out.push({ name, tags, score: matched * 1000 + score - name.length * 0.01 })
  }
  out.sort((a, b) => b.score - a.score || (a.name < b.name ? -1 : 1))
  return out.slice(0, limit)
}

/** Lime and white are for navy panels only; green, navy and charcoal for light backgrounds. */
export function iconColourAllowed(colour: IconColourName, onNavy: boolean): boolean {
  return onNavy ? colour === 'white' || colour === 'lime' : colour === 'green' || colour === 'navy' || colour === 'charcoal'
}
