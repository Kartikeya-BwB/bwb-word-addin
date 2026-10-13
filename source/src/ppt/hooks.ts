// React hooks for the PowerPoint pane: the live selection, and small per-person preferences kept in the browser.

import { useCallback, useEffect, useState } from 'react'
import type { SelectionState } from './selection.ts'
import { selectionTracker } from './selection.ts'

/** The selection, kept up to date while the component is mounted: it listens to selection changes meanwhile. */
export function useSelection(): SelectionState {
  const [state, setState] = useState(selectionTracker.current)
  useEffect(() => {
    const off = selectionTracker.subscribe(setState)
    void selectionTracker.start()
    return () => {
      off()
      void selectionTracker.stop()
    }
  }, [])
  return state
}

/** A remembered choice, such as what to line up with. Falls back to `initial` when storage is unavailable. */
export function usePref<T extends string | number | boolean>(key: string, initial: T, allowed?: readonly T[]): [T, (value: T) => void] {
  const [value, setValue] = useState<T>(() => {
    try {
      const raw = localStorage.getItem(key)
      const parsed: unknown = raw === null ? null : JSON.parse(raw)
      if (typeof parsed === typeof initial && (!allowed || allowed.includes(parsed as T))) return parsed as T
    } catch {
      // No storage here (a private window, a preview): keep the default.
    }
    return initial
  })
  const set = useCallback((next: T) => {
    setValue(next)
    try {
      localStorage.setItem(key, JSON.stringify(next))
    } catch {
      // As above: the choice lasts for this session only.
    }
  }, [key])
  return [value, set]
}
