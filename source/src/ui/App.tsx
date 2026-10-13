import { Suspense, useEffect, useMemo, useRef, useState } from 'react'
import type { KeyboardEvent } from 'react'
import { currentHost } from '../host.ts'
import { hasPowerPoint } from '../ppt/run.ts'
import { hasWord } from '../word/run.ts'
import { CommandSearch } from './CommandSearch.tsx'
import { ToastProvider } from './toast.tsx'
import type { TabDef } from './tabSets.tsx'
import { HOST_TAGLINE, tabsFor } from './tabSets.tsx'

const VERSION = '2.1'

/** Ribbon buttons open the pane at a tab, and optionally a section of it, through the URL hash (#insert/icons). */
function fromHash(tabs: TabDef[]): { tab: string; section: string | null } {
  const [tab = '', section = null] = window.location.hash.replace('#', '').split('/')
  const found = tabs.find((t) => t.id === tab) ?? tabs[0]
  return { tab: found?.id ?? '', section }
}

export function App() {
  const host = useMemo(currentHost, [])
  const tabs = useMemo(() => tabsFor(host), [host])
  const [tab, setTab] = useState(() => fromHash(tabs).tab)
  const [section, setSection] = useState<string | null>(() => fromHash(tabs).section)
  const tabRefs = useRef<Array<HTMLButtonElement | null>>([])
  const connected = host === 'powerpoint' ? hasPowerPoint() : hasWord()
  const current = tabs.find((t) => t.id === tab) ?? tabs[0]

  useEffect(() => {
    const onHash = (): void => {
      const next = fromHash(tabs)
      setTab(next.tab)
      setSection(next.section)
    }
    window.addEventListener('hashchange', onHash)
    return () => window.removeEventListener('hashchange', onHash)
  }, [tabs])

  const select = (id: string): void => {
    setTab(id)
    setSection(null)
    window.history.replaceState(null, '', `#${id}`)
  }

  const onKey = (e: KeyboardEvent<HTMLDivElement>): void => {
    const i = tabs.findIndex((t) => t.id === tab)
    const step = e.key === 'ArrowRight' ? 1 : e.key === 'ArrowLeft' ? -1 : 0
    if (!step) return
    const next = tabs[(i + step + tabs.length) % tabs.length]
    if (!next) return
    select(next.id)
    tabRefs.current[tabs.indexOf(next)]?.focus()
  }

  return (
    <ToastProvider>
      <div className={`shell host-${host}`}>
        <header className="masthead">
          <img src="./assets/brand/logo-white.png" alt="Bankers without Boundaries" />
          <div>
            <div className="masthead-title">Toolkit</div>
            <div className="masthead-sub">{HOST_TAGLINE[host]}</div>
          </div>
        </header>
        <div className="strip" aria-hidden><span /><span /><span /></div>
        <div className="tabs" role="tablist" aria-label="Toolkit sections" onKeyDown={onKey}>
          {tabs.map(({ id, label, Icon }, i) => (
            <button
              key={id}
              ref={(el) => { tabRefs.current[i] = el }}
              type="button"
              role="tab"
              id={`tab-${id}`}
              aria-selected={tab === id}
              aria-controls={`panel-${id}`}
              tabIndex={tab === id ? 0 : -1}
              className="tab"
              onClick={() => select(id)}
            >
              <Icon size={17} strokeWidth={tab === id ? 2.2 : 1.8} aria-hidden />
              {label}
            </button>
          ))}
        </div>
        <CommandSearch host={host} />
        <main className="panel" role="tabpanel" id={`panel-${tab}`} aria-labelledby={`tab-${tab}`}>
          {!connected && (
            <div className="notice warn">
              Preview mode: this pane is open in a browser, so buttons that change a {host === 'powerpoint' ? 'presentation' : 'document'} do
              nothing. Open it from {host === 'powerpoint' ? 'PowerPoint' : 'Word'} to use it.
            </div>
          )}
          {current && (
            <Suspense fallback={<p className="muted small">Loading…</p>}>
              <TabPanel key={current.id} def={current} section={section} />
            </Suspense>
          )}
        </main>
        <footer className="footer-note">
          <span>BwB Toolkit {VERSION}</span>
          <span>bwb.earth</span>
        </footer>
      </div>
    </ToastProvider>
  )
}

/** One tab's panel. It scrolls to the section the ribbon asked for once the tab's code has loaded and drawn. */
function TabPanel({ def, section }: { def: TabDef; section: string | null }) {
  const { Panel } = def
  useEffect(() => {
    if (section) document.getElementById(`section-${section}`)?.scrollIntoView({ block: 'start' })
  }, [section])
  return <Panel />
}
