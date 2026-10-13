import { useState } from 'react'
import { Check, Copy } from 'lucide-react'
import { Section, Divider } from './ui.jsx'

const PALETTE = [
  {
    group: 'Primary',
    colours: [
      { name: 'BwB Navy', hex: '#001522', cmyk: 'C100 M78 Y44 K79', use: 'Backgrounds, headings, body text' },
      { name: 'Lime', hex: '#BED62F', cmyk: 'C11 M0 Y93 K16', use: 'Accents, highlights, bullets' },
      { name: 'Forest Green', hex: '#328840', cmyk: 'C76 M15 Y100 K4', use: 'Data tables, numbered lists' },
      { name: 'Charcoal', hex: '#4A4A4C', cmyk: 'C66 M58 Y54 K32', use: 'Secondary body text, footers' },
    ],
  },
  {
    group: 'Extended',
    colours: [
      { name: 'White', hex: '#FFFFFF', cmyk: 'C0 M0 Y0 K0', use: 'Text on dark backgrounds' },
      { name: 'Light Grey', hex: '#F4F4F4', cmyk: 'C3 M2 Y2 K0', use: 'Page backgrounds, table fills' },
      { name: 'Mid Grey', hex: '#E0E0E0', cmyk: 'C11 M8 Y8 K0', use: 'Borders, dividers' },
      { name: 'Navy Mid', hex: '#0A2540', cmyk: 'C100 M78 Y30 K60', use: 'Deep navy for UI elements' },
    ],
  },
]

export default function ColoursPanel({ notify }) {
  const [copied, setCopied] = useState(null)

  const copy = (hex, name) => {
    navigator.clipboard
      .writeText(hex)
      .then(() => {
        setCopied(hex)
        notify(`Copied ${name}: ${hex}`)
        setTimeout(() => setCopied(null), 1800)
      })
      .catch(() => {
        notify('Clipboard not available', 'err')
      })
  }

  return (
    <div>
      {PALETTE.map(({ group, colours }) => (
        <Section key={group} title={group}>
          {colours.map((c) => (
            <div key={c.hex} style={s.row}>
              <div
                style={{
                  ...s.swatch,
                  background: c.hex,
                  border: c.hex === '#FFFFFF' ? '1px solid #e0e0e0' : 'none',
                }}
              />
              <div style={s.meta}>
                <div style={s.name}>{c.name}</div>
                <div style={s.hex}>{c.hex}</div>
                <div style={s.use}>{c.use}</div>
                <div style={s.cmyk}>{c.cmyk}</div>
              </div>
              <button
                onClick={() => copy(c.hex, c.name)}
                style={{ ...s.copyBtn, ...(copied === c.hex ? s.copyBtnDone : {}) }}
                title={`Copy ${c.hex}`}
              >
                {copied === c.hex ? <Check size={13} strokeWidth={2.5} /> : <Copy size={13} strokeWidth={1.8} />}
              </button>
            </div>
          ))}
        </Section>
      ))}

      <Divider />

      <Section title="Typography">
        <div style={s.fontRow}>
          <div style={{ ...s.fontSample, fontFamily: 'Candara, Georgia, serif' }}>Candara</div>
          <div style={s.fontMeta}>
            <div style={s.fontName}>Candara</div>
            <div style={s.fontUse}>Headings H1 – H3</div>
          </div>
        </div>
        <div style={s.fontRow}>
          <div style={{ ...s.fontSample, fontFamily: 'Aptos, system-ui, sans-serif' }}>Aptos</div>
          <div style={s.fontMeta}>
            <div style={s.fontName}>Aptos</div>
            <div style={s.fontUse}>Body, tables, captions</div>
          </div>
        </div>
      </Section>
    </div>
  )
}

const s = {
  row: {
    display: 'flex',
    alignItems: 'flex-start',
    gap: 10,
    padding: '8px 10px',
    background: '#f7f9fa',
    border: '1px solid #e8eef2',
    borderRadius: 6,
    marginBottom: 5,
  },
  swatch: { width: 34, height: 34, borderRadius: 4, flexShrink: 0 },
  meta: { flex: 1 },
  name: { fontWeight: 600, fontSize: 12, color: '#001522', lineHeight: 1.3 },
  hex: { fontSize: 11, color: '#001522', fontFamily: 'monospace', marginTop: 1, fontWeight: 600 },
  use: { fontSize: 10, color: '#6b7f8c', marginTop: 2, lineHeight: 1.4 },
  cmyk: { fontSize: 9.5, color: '#9eaab4', marginTop: 1, fontFamily: 'monospace' },
  copyBtn: {
    width: 28,
    height: 28,
    border: '1px solid #e8eef2',
    borderRadius: 5,
    background: '#fff',
    color: '#6b7f8c',
    display: 'flex',
    alignItems: 'center',
    justifyContent: 'center',
    cursor: 'pointer',
    flexShrink: 0,
    transition: 'all 0.15s',
    alignSelf: 'center',
  },
  copyBtnDone: { background: '#328840', color: '#fff', borderColor: '#328840' },
  fontRow: {
    display: 'flex',
    alignItems: 'center',
    gap: 10,
    padding: '8px 10px',
    background: '#f7f9fa',
    border: '1px solid #e8eef2',
    borderRadius: 6,
    marginBottom: 5,
  },
  fontSample: {
    width: 48,
    height: 34,
    display: 'flex',
    alignItems: 'center',
    justifyContent: 'center',
    background: '#001522',
    color: '#bed62f',
    borderRadius: 4,
    fontSize: 11,
    fontWeight: 700,
    flexShrink: 0,
  },
  fontMeta: { flex: 1 },
  fontName: { fontWeight: 600, fontSize: 12, color: '#001522' },
  fontUse: { fontSize: 10.5, color: '#6b7f8c', marginTop: 1 },
}
