import { runWord } from '../App.jsx'
import { Section, Divider } from './ui.jsx'

const applyListStyle = async (style, notify) => {
  try {
    await runWord(async (context) => {
      const paragraphs = context.document.getSelection().paragraphs
      paragraphs.load('items')
      await context.sync()
      for (const p of paragraphs.items) {
        p.style = style
      }
      await context.sync()
    })
    notify(`Applied: ${style}`)
  } catch (e) {
    notify(`Error: ${e.message}`, 'err')
  }
}

const BULLETS = [
  {
    style: 'BwB Bullet Lime',
    label: 'Lime Square',
    sub: 'Primary bullets',
    marker: '■',
    markerColor: '#bed62f',
    markerBg: '#001522',
  },
  {
    style: 'BwB Bullet Navy',
    label: 'Navy Circle',
    sub: 'Secondary / nested',
    marker: '●',
    markerColor: '#fff',
    markerBg: '#001522',
  },
  {
    style: 'BwB Bullet Green',
    label: 'Green Square',
    sub: 'Tables, callouts',
    marker: '■',
    markerColor: '#fff',
    markerBg: '#328840',
  },
  {
    style: 'BwB Bullet Arrow',
    label: 'Lime Arrow',
    sub: 'Action items, next steps',
    marker: '›',
    markerColor: '#bed62f',
    markerBg: '#001522',
    markerSize: 20,
  },
]

const NUMBERED = [
  {
    style: 'BwB Numbered Navy',
    label: 'Navy Numbered',
    sub: 'Ordered steps',
    marker: '1.',
    markerColor: '#fff',
    markerBg: '#001522',
    markerSize: 11,
  },
  {
    style: 'List Number',
    label: 'Standard Numbered',
    sub: 'Default Word list',
    marker: '1.',
    markerColor: '#328840',
    markerBg: '#e8f5ea',
    markerSize: 11,
  },
]

export default function ListsPanel({ notify }) {
  return (
    <div>
      <Section title="BwB Bullets">
        {BULLETS.map((b) => (
          <button key={b.style} onClick={() => applyListStyle(b.style, notify)} style={s.btn}>
            <div
              style={{
                ...s.marker,
                background: b.markerBg,
                color: b.markerColor,
                fontSize: b.markerSize || 14,
              }}
            >
              {b.marker}
            </div>
            <div style={s.meta}>
              <div style={s.label}>{b.label}</div>
              <div style={s.sub}>{b.sub}</div>
            </div>
            <div style={s.sampleLines}>
              <div style={{ ...s.sampleLine, width: 44 }} />
              <div style={{ ...s.sampleLine, width: 32 }} />
            </div>
          </button>
        ))}
      </Section>

      <Divider />

      <Section title="Numbered Lists">
        {NUMBERED.map((n) => (
          <button key={n.style} onClick={() => applyListStyle(n.style, notify)} style={s.btn}>
            <div
              style={{
                ...s.marker,
                background: n.markerBg,
                color: n.markerColor,
                fontSize: n.markerSize || 14,
                fontWeight: 700,
              }}
            >
              {n.marker}
            </div>
            <div style={s.meta}>
              <div style={s.label}>{n.label}</div>
              <div style={s.sub}>{n.sub}</div>
            </div>
          </button>
        ))}
      </Section>

      <Divider />

      <Section title="Tip">
        <div style={s.tip}>
          Select one or more paragraphs, then tap a list style. For sub-levels, select the line and use
          Tab in Word, or apply a 2nd-level style manually.
        </div>
      </Section>
    </div>
  )
}

const s = {
  btn: {
    display: 'flex',
    alignItems: 'center',
    gap: 10,
    width: '100%',
    padding: '8px 10px',
    background: '#f7f9fa',
    border: '1px solid #e8eef2',
    borderRadius: 6,
    marginBottom: 5,
    cursor: 'pointer',
    textAlign: 'left',
    fontFamily: 'inherit',
  },
  marker: {
    width: 34,
    height: 34,
    display: 'flex',
    alignItems: 'center',
    justifyContent: 'center',
    borderRadius: 4,
    fontWeight: 700,
    flexShrink: 0,
  },
  meta: { flex: 1 },
  label: { fontWeight: 600, fontSize: 12, color: '#001522' },
  sub: { fontSize: 10.5, color: '#6b7f8c', marginTop: 1 },
  sampleLines: { display: 'flex', flexDirection: 'column', gap: 4, justifyContent: 'center' },
  sampleLine: { height: 3, background: '#e0e0e0', borderRadius: 2 },
  tip: {
    background: '#f0f7f1',
    border: '1px solid #c5dfc8',
    borderRadius: 6,
    padding: '9px 11px',
    fontSize: 11,
    color: '#2a6b34',
    lineHeight: 1.5,
  },
}
