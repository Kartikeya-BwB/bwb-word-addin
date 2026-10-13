import { useEffect, useRef, useState } from 'react'
import type { KeyboardEvent } from 'react'
import { FilePlus2, Palette, PenLine, ShieldCheck, SquarePlus } from 'lucide-react'
import type { LucideIcon } from 'lucide-react'
import { hasWord } from '../word/run.ts'
import { ToastProvider } from './toast.tsx'
import { StartTab } from './tabs/StartTab.tsx'
import { WriteTab } from './tabs/WriteTab.tsx'
import { InsertTab } from './tabs/InsertTab.tsx'
import { CheckTab } from './tabs/CheckTab.tsx'
import { BrandTab } from './tabs/BrandTab.tsx'

export type TabId = 'start' | 'write' | 'insert' | 'check' | 'brand'

const TABS: Array<{ id: TabId; label: string; Icon: LucideIcon }> = [
  { id: 'start', label: 'Start', Icon: FilePlus2 },
  { id: 'write', label: 'Write', Icon: PenLine },
  { id: 'insert', label: 'Insert', Icon: SquarePlus },
  { id: 'check', label: 'Check', Icon: ShieldCheck },
  { id: 'brand', label: 'Brand', Icon: Palette },
]

const VERSION = '2.0'

/** Ribbon buttons open the pane at a tab, and optionally a section of it, through the URL hash (#insert/icons). */
function fromHash(): { tab: TabId; section: string | null } {
  const [tab = '', section = null] = window.location.hash.replace('#', '').split('/')
  const found = TABS.find((t) => t.id === tab)
  return { tab: found ? found.id : 'start', section }
}

export function App() {
  const [tab, setTab] = useState<TabId>(() => fromHash().tab)
  const [section, setSection] = useState<string | null>(() => fromHash().section)
  const tabRefs = useRef<Array<HTMLButtonElement | null>>([])

  useEffect(() => {
    const onHash = (): void => {
      const next = fromHash()
      setTab(next.tab)
      setSection(next.section)
    }
    window.addEventListener('hashchange', onHash)
    return () => window.removeEventListener('hashchange', onHash)
  }, [])

  useEffect(() => {
    if (section) document.getElementById(`section-${section}`)?.scrollIntoView({ block: 'start' })
  }, [tab, section])

  const select = (id: TabId): void => {
    setTab(id)
    setSection(null)
    window.history.replaceState(null, '', `#${id}`)
  }

  const onKey = (e: KeyboardEvent<HTMLDivElement>): void => {
    const i = TABS.findIndex((t) => t.id === tab)
    const step = e.key === 'ArrowRight' ? 1 : e.key === 'ArrowLeft' ? -1 : 0
    if (!step) return
    const next = TABS[(i + step + TABS.length) % TABS.length]
    if (!next) return
    select(next.id)
    tabRefs.current[TABS.indexOf(next)]?.focus()
  }

  return (
    <ToastProvider>
      <div className="shell">
        <header className="masthead">
          <img src="./assets/brand/logo-white.png" alt="Bankers without Boundaries" />
          <div>
            <div className="masthead-title">Toolkit</div>
            <div className="masthead-sub">Word documents in the BwB format</div>
          </div>
        </header>
        <div className="strip" aria-hidden><span /><span /><span /></div>
        <div className="tabs" role="tablist" aria-label="Toolkit sections" onKeyDown={onKey}>
          {TABS.map(({ id, label, Icon }, i) => (
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
        <main className="panel" role="tabpanel" id={`panel-${tab}`} aria-labelledby={`tab-${tab}`}>
          {!hasWord() && (
            <div className="notice warn">
              Preview mode: this pane is open in a browser, so buttons that change a document do nothing. Open it from Word to use it.
            </div>
          )}
          {tab === 'start' && <StartTab />}
          {tab === 'write' && <WriteTab />}
          {tab === 'insert' && <InsertTab />}
          {tab === 'check' && <CheckTab />}
          {tab === 'brand' && <BrandTab />}
        </main>
        <footer className="footer-note">
          <span>BwB Toolkit {VERSION}</span>
          <span>bwb.earth</span>
        </footer>
      </div>
    </ToastProvider>
  )
}
