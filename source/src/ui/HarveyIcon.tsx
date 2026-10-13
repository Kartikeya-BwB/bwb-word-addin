/**
 * A Harvey ball drawn in the current colour, for pane buttons. Fonts draw the ◔ ◑ ◕ glyphs unevenly, and not at all
 * where Segoe UI Symbol is missing, so the pane draws its own.
 */
export function HarveyIcon({ quarters, size = 22 }: { quarters: number; size?: number }) {
  const c = 12
  const r = 9.5
  const turn = Math.max(0, Math.min(4, quarters)) / 4
  const angle = turn * 2 * Math.PI
  const x = c + r * Math.sin(angle)
  const y = c - r * Math.cos(angle)
  return (
    <svg width={size} height={size} viewBox="0 0 24 24" aria-hidden focusable="false">
      <circle cx={c} cy={c} r={r} fill={turn === 1 ? 'currentColor' : 'none'} stroke="currentColor" strokeWidth={1.6} />
      {turn > 0 && turn < 1 && (
        <path d={`M${c} ${c} L${c} ${c - r} A${r} ${r} 0 ${turn > 0.5 ? 1 : 0} 1 ${x.toFixed(3)} ${y.toFixed(3)} Z`} fill="currentColor" />
      )}
    </svg>
  )
}
