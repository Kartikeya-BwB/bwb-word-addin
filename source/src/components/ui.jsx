export function Section({ title, children }) {
  return (
    <div style={s.section}>
      <div style={s.sectionTitle}>{title}</div>
      {children}
    </div>
  )
}

export function Divider() {
  return <div style={s.divider} />
}

// Only section, sectionTitle and divider are used; the rest shipped unused in the original build.
const s = {
  section: { marginBottom: 18 },
  sectionTitle: {
    fontSize: 9.5,
    fontWeight: 700,
    letterSpacing: 0.8,
    textTransform: 'uppercase',
    color: '#8aabbf',
    marginBottom: 8,
    paddingBottom: 4,
    borderBottom: '1px solid #e8eef2',
  },
  styleBtn: {
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
    transition: 'all 0.12s',
    textAlign: 'left',
  },
  preview: {
    width: 32,
    height: 32,
    background: '#001522',
    color: '#fff',
    display: 'flex',
    alignItems: 'center',
    justifyContent: 'center',
    borderRadius: 4,
    fontWeight: 700,
    fontSize: 14,
    flexShrink: 0,
  },
  styleMeta: { flex: 1 },
  styleLabel: { fontWeight: 600, fontSize: 12, color: '#001522', lineHeight: 1.3 },
  styleSub: { fontSize: 10.5, color: '#6b7f8c', marginTop: 1 },
  iconBtn: {
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
    transition: 'all 0.12s',
    textAlign: 'left',
  },
  iconWrap: {
    width: 32,
    height: 32,
    borderRadius: 6,
    display: 'flex',
    alignItems: 'center',
    justifyContent: 'center',
    flexShrink: 0,
  },
  iconMeta: { flex: 1 },
  iconLabel: { fontWeight: 600, fontSize: 12, color: '#001522' },
  iconSub: { fontSize: 10.5, color: '#6b7f8c', marginTop: 1 },
  chevron: { color: '#bed62f', fontSize: 18, fontWeight: 700, lineHeight: 1 },
  divider: { height: 1, background: '#e8eef2', margin: '12px 0' },
}
