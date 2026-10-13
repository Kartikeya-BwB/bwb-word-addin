// Which shapes are selected, in the order they joined the selection, and the anchor (think-cell's key object).
// PowerPoint does not report the order of selection, so while the pane is open it listens to DocumentSelectionChanged
// and diffs the selected ids. The state goes to localStorage too, so the ribbon's commands, which run in a page of
// their own, use the same anchor and order.

import { UserFacingError } from '../word/run.ts'
import { hasPowerPoint, runPpt, supportsPpt } from './run.ts'

export interface SelectedShape {
  id: string
  name: string
}

export interface SelectionRead {
  /** The active slide. */
  slideId: string | null
  shapes: SelectedShape[]
}

export interface Anchor {
  id: string
  name: string
  slideId: string
}

export interface SelectionState {
  slideId: string | null
  /** The selected shapes' ids, in the order they joined the selection as far as the pane saw. */
  order: string[]
  names: Record<string, string>
  /** The anchor set by hand. It may be outside the selection; it applies on its own slide only. */
  anchor: Anchor | null
  /** When the state was last updated (milliseconds since 1970). */
  at: number
}

const KEY = 'bwb.ppt.selection.v1'
/** The ribbon ignores state the pane wrote longer ago than this. */
const STALE_MS = 12 * 60 * 60 * 1000

export const emptyState = (): SelectionState => ({ slideId: null, order: [], names: {}, anchor: null, at: 0 })

/** The order after a selection change: shapes still selected keep their order, new ones follow in the order given. */
export function nextOrder(previous: readonly string[], current: readonly string[]): string[] {
  const now = new Set(current)
  const kept = previous.filter((id) => now.has(id))
  const known = new Set(kept)
  return [...kept, ...current.filter((id) => !known.has(id))]
}

function storage(): Storage | null {
  try {
    return typeof localStorage === 'undefined' ? null : localStorage
  } catch {
    return null
  }
}

function isState(value: unknown): value is SelectionState {
  const v = value as Partial<SelectionState> | null
  return Boolean(v) && Array.isArray(v?.order) && typeof v?.names === 'object' && typeof v?.at === 'number'
}

/** The state the pane last saved, or an empty one when there is none or it is stale. */
export function loadState(now = Date.now()): SelectionState {
  try {
    const raw = storage()?.getItem(KEY)
    const parsed: unknown = raw ? JSON.parse(raw) : null
    return isState(parsed) && now - parsed.at < STALE_MS ? parsed : emptyState()
  } catch {
    return emptyState()
  }
}

function saveState(state: SelectionState): void {
  try {
    storage()?.setItem(KEY, JSON.stringify(state))
  } catch {
    // Private windows and locked-down browsers refuse storage; the pane still works, the ribbon falls back.
  }
}

/** Reads the active slide and the selected shapes in one sync. */
export async function readSelection(): Promise<SelectionRead> {
  return runPpt(async (context) => {
    const shapes = context.presentation.getSelectedShapes()
    const slides = context.presentation.getSelectedSlides()
    shapes.load('items/id,items/name')
    slides.load('items/id')
    await context.sync()
    return { slideId: slides.items[0]?.id ?? null, shapes: shapes.items.map((s) => ({ id: s.id, name: s.name })) }
  })
}

export class SelectionTracker {
  private state: SelectionState
  private readonly listeners = new Set<(state: SelectionState) => void>()
  private started = false
  private running: Promise<void> | null = null
  private again = false
  private readonly onChange = (): Promise<void> => this.refresh()

  constructor(private readonly read: () => Promise<SelectionRead> = readSelection, private readonly clock: () => number = Date.now) {
    this.state = loadState(clock())
  }

  get current(): SelectionState {
    return this.state
  }

  get listening(): boolean {
    return this.started
  }

  subscribe(listener: (state: SelectionState) => void): () => void {
    this.listeners.add(listener)
    return () => this.listeners.delete(listener)
  }

  /** Listens to selection changes (PowerPointApi 1.5) and reads the selection once. */
  async start(): Promise<void> {
    if (this.started || !hasPowerPoint() || !supportsPpt('1.5')) return
    this.started = true
    await new Promise<void>((resolve) => {
      Office.context.document.addHandlerAsync(Office.EventType.DocumentSelectionChanged, this.onChange, () => resolve())
    })
    await this.refresh()
  }

  async stop(): Promise<void> {
    if (!this.started) return
    this.started = false
    await new Promise<void>((resolve) => {
      Office.context.document.removeHandlerAsync(Office.EventType.DocumentSelectionChanged, { handler: this.onChange }, () => resolve())
    })
  }

  /** Reads the selection. Changes that arrive during a read are read once more afterwards, so none is lost. */
  refresh(): Promise<void> {
    if (this.running) {
      this.again = true
      return this.running
    }
    this.running = (async () => {
      try {
        do {
          this.again = false
          this.apply(await this.read())
        } while (this.again)
      } catch {
        // PowerPoint can refuse a read while a slide is changing; the next selection change reads again.
      } finally {
        this.running = null
      }
    })()
    return this.running
  }

  /** Folds a read into the state. On the same slide, shapes still selected keep their order; a new slide starts afresh. */
  apply(read: SelectionRead): void {
    const ids = read.shapes.map((s) => s.id)
    const order = nextOrder(read.slideId === this.state.slideId ? this.state.order : [], ids)
    const names = Object.fromEntries(read.shapes.map((s) => [s.id, s.name]))
    this.set({ ...this.state, slideId: read.slideId, order, names })
  }

  /** Makes the one selected shape the anchor; returns a message for the pane. */
  async setAnchorFromSelection(): Promise<string> {
    const read = await this.read()
    this.apply(read)
    const [only] = read.shapes
    if (read.shapes.length !== 1 || !only || !read.slideId) throw new UserFacingError('Select one shape to make it the anchor.')
    this.set({ ...this.state, anchor: { id: only.id, name: only.name, slideId: read.slideId } })
    return `“${only.name}” is the anchor. Alignment to the anchor and Same size use it.`
  }

  clearAnchor(): void {
    this.set({ ...this.state, anchor: null })
  }

  private set(next: SelectionState): void {
    this.state = { ...next, at: this.clock() }
    saveState(this.state)
    this.listeners.forEach((listener) => listener(this.state))
  }
}

/** The tracker the pane starts. In the ribbon's page it never starts, and the saved state stands in for it. */
export const selectionTracker = new SelectionTracker()

export interface AnchorHint {
  /** The anchor: the one set by hand on this slide, else the first shape selected, else none (the largest is used). */
  anchorId: string | null
  /** Set by hand: it may be outside the selection. */
  explicit: boolean
  /** Ids of the selected shapes in the order they were picked, as far as that is known. */
  order: string[]
}

/** The anchor and selection order for the selection on a slide, from the pane's live state or the saved one. */
export function anchorHint(slideId: string | null, selected: readonly string[], state: SelectionState = currentState()): AnchorHint {
  const order = state.slideId === slideId ? state.order.filter((id) => selected.includes(id)) : []
  if (state.anchor && state.anchor.slideId === slideId) return { anchorId: state.anchor.id, explicit: true, order }
  return { anchorId: order[0] ?? null, explicit: false, order }
}

/** The fresher of this page's state and the saved one: the ribbon's page sees what the pane saved. */
export function currentState(): SelectionState {
  const saved = loadState()
  return selectionTracker.current.at >= saved.at ? selectionTracker.current : saved
}
