import { useState, useCallback } from 'react'
import { Type, List, SquarePlus, Palette, CircleCheckBig, CircleAlert } from 'lucide-react'
import StylesPanel from './components/StylesPanel.jsx'
import ListsPanel from './components/ListsPanel.jsx'
import InsertPanel from './components/InsertPanel.jsx'
import ColoursPanel from './components/ColoursPanel.jsx'

const TABS = [
  { id: 'styles', label: 'Styles', Icon: Type },
  { id: 'lists', label: 'Lists', Icon: List },
  { id: 'insert', label: 'Insert', Icon: SquarePlus },
  { id: 'colours', label: 'Colours', Icon: Palette },
]

// Runs a Word.run batch, or does nothing when opened outside Word (e.g. `npm run dev` in a browser).
export const runWord = async (fn) => {
  if (typeof Word === 'undefined') {
    console.warn('Word API not available (dev mode)')
    return null
  }
  return Word.run(fn)
}

export default function App() {
  const [tab, setTab] = useState('styles')
  const [toast, setToast] = useState(null)

  const notify = useCallback((msg, type = 'ok') => {
    setToast({ msg, type })
    setTimeout(() => setToast(null), 2800)
  }, [])

  return (
    <div style={s.shell}>
      <header style={s.header}>
        <div style={s.headerInner}>
          <div style={s.logoMark}>BwB</div>
          <div>
            <div style={s.logoTitle}>Bankers without Boundaries</div>
            <div style={s.logoSub}>Document Assistant</div>
          </div>
        </div>
        <div style={s.limeLine} />
      </header>

      <nav style={s.tabBar}>
        {TABS.map(({ id, label, Icon }) => (
          <button
            key={id}
            onClick={() => setTab(id)}
            style={{ ...s.tabBtn, ...(tab === id ? s.tabActive : {}) }}
          >
            <Icon size={15} strokeWidth={tab === id ? 2.2 : 1.8} />
            <span style={s.tabLabel}>{label}</span>
          </button>
        ))}
      </nav>

      <main style={s.panel}>
        {tab === 'styles' && <StylesPanel notify={notify} />}
        {tab === 'lists' && <ListsPanel notify={notify} />}
        {tab === 'insert' && <InsertPanel notify={notify} />}
        {tab === 'colours' && <ColoursPanel notify={notify} />}
      </main>

      {toast && (
        <div style={{ ...s.toast, ...(toast.type === 'err' ? s.toastErr : {}) }}>
          {toast.type === 'ok' ? <CircleCheckBig size={14} /> : <CircleAlert size={14} />}
          <span>{toast.msg}</span>
        </div>
      )}

      <footer style={s.footer}>bwb.earth</footer>
    </div>
  )
}

const s = {
  shell: { display: 'flex', flexDirection: 'column', height: '100vh', background: '#fff', overflow: 'hidden' },
  header: { background: '#001522', padding: '14px 16px 0', flexShrink: 0 },
  headerInner: { display: 'flex', alignItems: 'center', gap: 10, paddingBottom: 12 },
  logoMark: {
    width: 36,
    height: 36,
    background: '#bed62f',
    color: '#001522',
    fontWeight: 700,
    fontSize: 13,
    display: 'flex',
    alignItems: 'center',
    justifyContent: 'center',
    borderRadius: 4,
    letterSpacing: 0.5,
    flexShrink: 0,
  },
  logoTitle: { color: '#fff', fontWeight: 700, fontSize: 11.5, lineHeight: 1.2, letterSpacing: 0.2 },
  logoSub: { color: '#bed62f', fontSize: 10, fontWeight: 300, marginTop: 1, letterSpacing: 0.3 },
  limeLine: { height: 2, background: '#bed62f', marginLeft: -16, marginRight: -16 },
  tabBar: { display: 'flex', background: '#0a2540', flexShrink: 0 },
  tabBtn: {
    flex: 1,
    display: 'flex',
    flexDirection: 'column',
    alignItems: 'center',
    gap: 3,
    padding: '9px 4px',
    background: 'transparent',
    color: '#8aabbf',
    fontSize: 10,
    transition: 'all 0.15s',
    borderBottom: '2px solid transparent',
    cursor: 'pointer',
    border: 'none',
  },
  tabActive: { color: '#bed62f', borderBottom: '2px solid #bed62f', background: 'rgba(190,214,47,0.06)' },
  tabLabel: { fontSize: 9.5, letterSpacing: 0.3, textTransform: 'uppercase', fontWeight: 600 },
  panel: { flex: 1, overflowY: 'auto', padding: '14px 14px 8px' },
  toast: {
    position: 'fixed',
    bottom: 30,
    left: '50%',
    transform: 'translateX(-50%)',
    background: '#328840',
    color: '#fff',
    padding: '8px 14px',
    borderRadius: 6,
    display: 'flex',
    alignItems: 'center',
    gap: 6,
    fontSize: 12,
    fontWeight: 600,
    boxShadow: '0 4px 16px rgba(0,0,0,0.2)',
    whiteSpace: 'nowrap',
    zIndex: 999,
  },
  toastErr: { background: '#c0392b' },
  footer: {
    textAlign: 'center',
    padding: '6px',
    fontSize: 10,
    color: '#bed62f',
    background: '#001522',
    letterSpacing: 0.5,
    flexShrink: 0,
  },
}
