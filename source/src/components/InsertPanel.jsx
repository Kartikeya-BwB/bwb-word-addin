import { Table2, Square, Minus, TextAlignStart } from 'lucide-react'
import { runWord } from '../App.jsx'
import { Section, Divider } from './ui.jsx'

const insertTable = async (rows, cols, styleName, notify, label) => {
  try {
    await runWord(async (context) => {
      const table = context.document.getSelection().getRange().insertTable(rows, cols, 'After', [])
      table.style = styleName
      table.styleBuiltIn = Word.Style.tableGrid
      const headerRow = table.rows.getFirst()
      headerRow.load('cells')
      await context.sync()
      for (const cell of headerRow.cells.items) {
        cell.shadingColor = '001522'
        cell.body.paragraphs.getFirst().load('text')
        await context.sync()
      }
      await context.sync()
    })
    notify(`${label} table inserted`)
  } catch {
    // Fall back to a plain table if styling fails.
    try {
      await runWord(async (context) => {
        context.document.getSelection().getRange().insertTable(rows, cols, 'After', [])
        await context.sync()
      })
      notify(`${label} table inserted (apply BwB style manually)`)
    } catch (e) {
      notify(`Error: ${e.message}`, 'err')
    }
  }
}

const insertCallout = async (notify) => {
  try {
    await runWord(async (context) => {
      const p = context.document
        .getSelection()
        .getRange()
        .insertParagraph('Key insight: Replace this text with your callout message.', 'After')
      p.style = 'Normal'
      p.leftIndent = 0
      await context.sync()
    })
    notify('Callout inserted — apply BwB Highlight table style')
  } catch (e) {
    notify(`Error: ${e.message}`, 'err')
  }
}

const insertPageBreak = async (notify) => {
  try {
    await runWord(async (context) => {
      context.document.getSelection().getRange().insertBreak(Word.BreakType.page, 'After')
      await context.sync()
    })
    notify('Page break inserted')
  } catch (e) {
    notify(`Error: ${e.message}`, 'err')
  }
}

const ELEMENTS = [
  {
    Icon: Table2,
    label: 'BwB Data Table',
    sub: '3 × 4 with navy header',
    accent: '#001522',
    action: (notify) => insertTable(4, 3, 'TableGrid', notify, 'Data'),
  },
  {
    Icon: Table2,
    label: 'Two-Column Layout',
    sub: '1 row × 2 cols',
    accent: '#328840',
    action: (notify) => insertTable(1, 2, 'TableGrid', notify, 'Two-column'),
  },
  {
    Icon: Square,
    label: 'Highlight Callout',
    sub: 'Lime-bordered insight box',
    accent: '#bed62f',
    accentText: '#001522',
    action: (notify) => insertCallout(notify),
  },
  {
    Icon: Minus,
    label: 'Page Break',
    sub: 'Start a new page',
    accent: '#4a4a4c',
    action: (notify) => insertPageBreak(notify),
  },
  {
    Icon: TextAlignStart,
    label: 'Section Divider',
    sub: 'Navy panel — copy from template',
    accent: '#001522',
    action: (notify) => notify('Copy the navy panel from the template and paste it here', 'err'),
  },
]

export default function InsertPanel({ notify }) {
  return (
    <div>
      <Section title="Insert Elements">
        {ELEMENTS.map(({ Icon, label, sub, accent, accentText, action }) => (
          <button key={label} onClick={() => action(notify)} style={s.btn}>
            <div style={{ ...s.iconWrap, background: accent || '#001522' }}>
              <Icon size={16} color={accentText || '#fff'} strokeWidth={1.8} />
            </div>
            <div style={s.meta}>
              <div style={s.label}>{label}</div>
              <div style={s.sub}>{sub}</div>
            </div>
            <div style={s.chevron}>›</div>
          </button>
        ))}
      </Section>

      <Divider />

      <Section title="BwB Table Styles">
        <div style={s.tip}>
          After inserting a table, right-click it and choose <strong>Table Style</strong> to apply one of:
          <div style={s.styleList}>
            <span style={s.styleTag}>BwBStandard</span>
            <span style={s.styleTag}>BwBComparison</span>
            <span style={s.styleTag}>BwBSnapshot</span>
            <span style={s.styleTag}>BwBHighlight</span>
          </div>
        </div>
      </Section>

      <Divider />

      <Section title="Tip">
        <div style={s.tip}>
          Place your cursor where you want the element, then tap an option above. Tables are inserted
          with a standard grid — apply a BwB table style to brand them.
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
    padding: '9px 10px',
    background: '#f7f9fa',
    border: '1px solid #e8eef2',
    borderRadius: 6,
    marginBottom: 5,
    cursor: 'pointer',
    textAlign: 'left',
    fontFamily: 'inherit',
  },
  iconWrap: {
    width: 34,
    height: 34,
    borderRadius: 6,
    display: 'flex',
    alignItems: 'center',
    justifyContent: 'center',
    flexShrink: 0,
  },
  meta: { flex: 1 },
  label: { fontWeight: 600, fontSize: 12, color: '#001522' },
  sub: { fontSize: 10.5, color: '#6b7f8c', marginTop: 1 },
  chevron: { color: '#bed62f', fontSize: 18, fontWeight: 700, lineHeight: 1 },
  tip: {
    background: '#f0f7f1',
    border: '1px solid #c5dfc8',
    borderRadius: 6,
    padding: '9px 11px',
    fontSize: 11,
    color: '#2a6b34',
    lineHeight: 1.6,
  },
  styleList: { display: 'flex', flexWrap: 'wrap', gap: 4, marginTop: 6 },
  styleTag: {
    background: '#001522',
    color: '#bed62f',
    fontSize: 9.5,
    fontWeight: 600,
    padding: '2px 7px',
    borderRadius: 4,
    letterSpacing: 0.3,
  },
}
