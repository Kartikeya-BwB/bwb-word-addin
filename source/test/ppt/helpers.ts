// Shared helpers for the PowerPoint tests.

import { readFileSync } from 'node:fs'
import { join } from 'node:path'
import type { IconIndex } from '../../src/brand/icons.ts'

/** A Storage kept in memory, for stubbing localStorage. */
export function memoryStorage(): Storage {
  const data = new Map<string, string>()
  return {
    get length() { return data.size },
    clear: () => data.clear(),
    getItem: (key: string) => data.get(key) ?? null,
    key: (index: number) => [...data.keys()][index] ?? null,
    removeItem: (key: string) => void data.delete(key),
    setItem: (key: string, value: string) => void data.set(key, String(value)),
  }
}

let index: IconIndex | undefined

/** The bundled Lucide index, read from disk once. */
export function iconIndex(): IconIndex {
  index ??= JSON.parse(readFileSync(join(import.meta.dirname, '../../public/assets/lucide/index.json'), 'utf8')) as IconIndex
  return index
}

/** A fetch that serves the icon index, as the pane's own server does. */
export const iconFetch = async (url: string | URL | Request): Promise<Response> => {
  if (String(url).endsWith('lucide/index.json')) return new Response(JSON.stringify(iconIndex()), { status: 200 })
  return new Response('Not found', { status: 404 })
}
