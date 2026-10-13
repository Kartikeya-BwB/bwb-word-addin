import { runWord } from '../App.jsx'
import { Section, Divider } from './ui.jsx'

const applyStyle = async (styleName, notify) => {
  try {
    await runWord(async (context) => {
      const paragraphs = context.document.getSelection().paragraphs
      paragraphs.load('items')
      await context.sync()
      for (const p of paragraphs.items) {
        p.style = styleName
      }
      await context.sync()
    })
    notify(`Applied: ${styleName}`)
  } catch (e) {
    notify(`Error: ${e.message}`, 'err')
  }
}

const HEADINGS = [
  {
    name: 'heading 1',
    label: 'Heading 1',
    sub: 'Candara Bold · 28pt · Navy',
    previewStyle: { fontSize: 18, fontFamily: 'Georgia, serif', fontWeight: 700 },
  },
  {
    name: 'heading 2',
    label: 'Heading 2',
    sub: 'Candara Bold · 20pt · Navy',
    previewStyle: { fontSize: 15, fontFamily: 'Georgia, serif', fontWeight: 700 },
  },
  {
    name: 'heading 3',
    label: 'Heading 3',
    sub: 'Candara Bold · 14pt · Navy',
    previewStyle: { fontSize: 12, fontFamily: 'Georgia, serif', fontWeight: 700 },
  },
  {
    name: 'heading 4',
    label: 'Heading 4',
    sub: 'Aptos Bold Italic · 12pt',
    previewStyle: { fontSize: 11, fontStyle: 'italic', fontWeight: 700 },
  },
]

const BODY_STYLES = [
  { name: 'Normal', label: 'Body Text', sub: 'Aptos · 11pt', badge: 'Aa', badgeBg: '#4a4a4c' },
  { name: 'Caption', label: 'Caption', sub: 'Aptos · 9pt · Italic', badge: 'Cc', badgeBg: '#6b7f8c' },
  { name: 'Strong1', label: 'Strong (BwB)', sub: 'Aptos Bold', badge: 'B', badgeBg: '#001522' },
  {
    name: 'Title',
    label: 'Title',
    sub: 'Candara · 28pt · White on navy',
    badge: 'T',
    badgeBg: '#bed62f',
    badgeColor: '#001522',
  },
]

export default function StylesPanel({ notify }) {
  return (
    <div>
      <Section title="Headings">
        {HEADINGS.map((h) => (
          <button key={h.name} onClick={() => applyStyle(h.name, notify)} style={s.btn}>
            <div style={{ ...s.preview, ...h.previewStyle }}>H</div>
            <div style={s.meta}>
              <div style={s.label}>{h.label}</div>
              <div style={s.sub}>{h.sub}</div>
            </div>
            <div style={s.pill}>{h.name.split(' ')[1]}</div>
          </button>
        ))}
      </Section>

      <Divider />

      <Section title="Body & Text">
        {BODY_STYLES.map((b) => (
          <button key={b.name} onClick={() => applyStyle(b.name, notify)} style={s.btn}>
            <div
              style={{
                ...s.badge,
                background: b.badgeBg || '#001522',
                color: b.badgeColor || '#fff',
              }}
            >
              {b.badge}
            </div>
            <div style={s.meta}>
              <div style={s.label}>{b.label}</div>
              <div style={s.sub}>{b.sub}</div>
            </div>
          </button>
        ))}
      </Section>

      <Divider />

      <Section title="Tip">
        <div style={s.tip}>
          Select any text or place cursor in a paragraph, then tap a style above to apply it instantly.
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
  preview: {
    width: 34,
    height: 34,
    background: '#001522',
    color: '#fff',
    display: 'flex',
    alignItems: 'center',
    justifyContent: 'center',
    borderRadius: 4,
    fontWeight: 700,
    flexShrink: 0,
    fontFamily: 'Georgia, serif',
  },
  badge: {
    width: 34,
    height: 34,
    display: 'flex',
    alignItems: 'center',
    justifyContent: 'center',
    borderRadius: 4,
    fontWeight: 700,
    fontSize: 14,
    flexShrink: 0,
  },
  meta: { flex: 1 },
  label: { fontWeight: 600, fontSize: 12, color: '#001522', lineHeight: 1.3 },
  sub: { fontSize: 10.5, color: '#6b7f8c', marginTop: 1 },
  pill: {
    background: '#bed62f',
    color: '#001522',
    fontWeight: 700,
    fontSize: 10,
    padding: '2px 7px',
    borderRadius: 10,
  },
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
