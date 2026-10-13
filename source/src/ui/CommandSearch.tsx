import { useEffect, useMemo, useRef, useState } from 'react'
import type { KeyboardEvent } from 'react'
import { Search } from 'lucide-react'
import type { Tool } from '../commands/catalogue.ts'
import { searchTools, toolsFor } from '../commands/catalogue.ts'
import { HANDLERS } from '../commands/handlers.ts'
import type { Host } from '../host.ts'
import { useAction } from './useAction.ts'

/** "Find a tool": every ribbon tool of this app, searchable by what it does. Press / to jump here. */
export function CommandSearch({ host }: { host: Host }) {
  const { busy, run } = useAction()
  const tools = useMemo(() => toolsFor(host), [host])
  const [query, setQuery] = useState('')
  const [active, setActive] = useState(0)
  const input = useRef<HTMLInputElement>(null)
  const hits = useMemo(() => searchTools(tools, query), [tools, query])

  useEffect(() => {
    const onKey = (e: globalThis.KeyboardEvent): void => {
      const target = e.target as HTMLElement | null
      if (e.key === '/' && target && !['INPUT', 'TEXTAREA', 'SELECT'].includes(target.tagName)) {
        e.preventDefault()
        input.current?.focus()
      }
    }
    window.addEventListener('keydown', onKey)
    return () => window.removeEventListener('keydown', onKey)
  }, [])

  const choose = (tool: Tool): void => {
    setQuery('')
    if (tool.pane) {
      window.location.hash = `#${tool.pane}`
      return
    }
    const handler = tool.fn ? HANDLERS[tool.fn] : undefined
    if (handler) void run(`tool:${tool.id}`, async () => {
      await handler()
      return `${tool.label}: done.`
    })
  }

  const onKeyDown = (e: KeyboardEvent<HTMLInputElement>): void => {
    if (e.key === 'ArrowDown') setActive((i) => Math.min(i + 1, hits.length - 1))
    else if (e.key === 'ArrowUp') setActive((i) => Math.max(i - 1, 0))
    else if (e.key === 'Escape') setQuery('')
    else if (e.key === 'Enter') {
      const tool = hits[active]
      if (tool) choose(tool)
    } else return
    e.preventDefault()
  }

  return (
    <div className="command-search" role="search">
      <Search size={14} aria-hidden className="command-search-icon" />
      <input ref={input} type="search" value={query} placeholder="Find a tool (press /)" aria-label="Find a tool"
        aria-expanded={hits.length > 0} aria-controls="command-results" role="combobox" aria-autocomplete="list"
        onChange={(e) => { setQuery(e.target.value); setActive(0) }} onKeyDown={onKeyDown} />
      {hits.length > 0 && (
        <ul id="command-results" className="command-results" role="listbox">
          {hits.map((tool, i) => (
            <li key={tool.id} role="option" aria-selected={i === active}>
              <button type="button" disabled={busy !== null} onMouseEnter={() => setActive(i)} onClick={() => choose(tool)}>
                <span className="command-label">{tool.label}</span>
                <span className="command-path">{tool.path}{tool.pane ? ' · opens in the pane' : ''}</span>
              </button>
            </li>
          ))}
        </ul>
      )}
    </div>
  )
}
