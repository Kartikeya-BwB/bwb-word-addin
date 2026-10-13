// The last ten charts inserted, kept in the browser's local storage so people can reopen, tweak and reinsert them.
// Storage can be missing or full (private windows, locked-down Office installs); then the list is simply empty.

import { CHART_TYPES } from '../../charts/spec.ts'
import type { Host } from '../../host.ts'
import type { BuilderState } from './builderState.ts'
import { initialState } from './builderState.ts'

export const RECENT_KEY = 'bwb-toolkit.recent-charts'
export const MAX_RECENT = 10

export interface RecentChart {
  id: string
  /** ISO time it was inserted. */
  savedAt: string
  state: BuilderState
}

type Store = Pick<Storage, 'getItem' | 'setItem'>

function defaultStore(): Store | null {
  try {
    return typeof localStorage === 'undefined' ? null : localStorage
  } catch {
    return null
  }
}

const isGrid = (value: unknown): value is string[][] =>
  Array.isArray(value) && value.length > 0 && value.every((row) => Array.isArray(row) && row.every((c) => typeof c === 'string'))

/** A stored state made whole: fields added since it was saved take their defaults, unknown types are refused. */
function revive(value: unknown, host: Host): BuilderState | null {
  if (!value || typeof value !== 'object') return null
  const stored = value as Partial<BuilderState>
  if (!CHART_TYPES.some((t) => t.type === stored.type) || !isGrid(stored.grid)) return null
  const base = initialState(host, stored.type)
  const merged: Record<string, unknown> = { ...base }
  for (const [key, fallback] of Object.entries(base)) {
    const v = (stored as Record<string, unknown>)[key]
    if (v !== undefined && (typeof v === typeof fallback || (Array.isArray(fallback) && Array.isArray(v)))) merged[key] = v
  }
  return merged as unknown as BuilderState
}

export function loadRecent(host: Host, store: Store | null = defaultStore()): RecentChart[] {
  try {
    const raw = store?.getItem(RECENT_KEY)
    const parsed: unknown = raw ? JSON.parse(raw) : []
    if (!Array.isArray(parsed)) return []
    const out: RecentChart[] = []
    for (const item of parsed as Array<Partial<RecentChart>>) {
      const state = revive(item?.state, host)
      if (state && typeof item.id === 'string' && typeof item.savedAt === 'string') out.push({ id: item.id, savedAt: item.savedAt, state })
    }
    return out.slice(0, MAX_RECENT)
  } catch {
    return []
  }
}

/** Adds a chart to the top of the list (replacing an identical one) and returns the new list. */
export function saveRecent(state: BuilderState, host: Host, store: Store | null = defaultStore(), now: Date = new Date()): RecentChart[] {
  const current = loadRecent(host, store)
  const key = JSON.stringify({ ...state, example: false })
  const kept = current.filter((r) => JSON.stringify({ ...r.state, example: false }) !== key)
  const entry: RecentChart = { id: `${now.getTime().toString(36)}${Math.random().toString(36).slice(2, 6)}`, savedAt: now.toISOString(), state }
  const next = [entry, ...kept].slice(0, MAX_RECENT)
  try {
    store?.setItem(RECENT_KEY, JSON.stringify(next))
  } catch {
    // Full or blocked storage: the chart is still inserted, it just is not remembered.
  }
  return next
}

/** "Funding by year" or, untitled, the type's name. */
export function recentLabel(chart: RecentChart): string {
  return chart.state.title.trim() || (CHART_TYPES.find((t) => t.type === chart.state.type)?.name ?? 'Chart')
}
