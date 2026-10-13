// A stand-in for PowerPoint's editing window in the demo: slide thumbnails on the left and the active slide in the
// middle, drawn from the fake presentation. Clicking a shape selects it (Shift or Ctrl adds to the selection) and
// fires DocumentSelectionChanged, as PowerPoint does, so the pane works on what is selected here.

import type { MouseEvent, ReactNode } from 'react'
import type { FakePresentation, FakeRun, FakeShape, FakeSlide } from '../fakePpt/index.ts'
import { activeSlide, boundsOf, clickShapes, fireSelectionChanged, selectSlides } from '../fakePpt/index.ts'

/** Pictures the demo knows how to draw: what setSelectedDataAsync was given, or a library slide's thumbnail. */
export const pictureSources = new Map<string, string>()

function pictureUrl(model: FakePresentation, shape: FakeShape): string | null {
  const known = pictureSources.get(shape.id)
  if (known) return known
  const insert = model.dataInserts.find((d) => d.shapeId === shape.id)
  if (!insert) return null
  if (insert.coercionType === 'xmlSvg') return `data:image/svg+xml;charset=utf-8,${encodeURIComponent(insert.data)}`
  return `data:image/${insert.data.startsWith('/9j/') ? 'jpeg' : 'png'};base64,${insert.data}`
}

const fillOf = (s: FakeShape): string => (s.fill.type === 'Solid' ? s.fill.foregroundColor : 'none')
const strokeOf = (s: FakeShape): { stroke: string; strokeWidth: number } =>
  s.line.visible ? { stroke: s.line.color, strokeWidth: Math.max(0.5, s.line.weight) } : { stroke: 'none', strokeWidth: 0 }

function outline(s: FakeShape): ReactNode {
  const { left: x, top: y, width: w, height: h } = s
  const paint = { fill: fillOf(s), ...strokeOf(s) }
  switch (s.geometricShapeType) {
    case 'Ellipse':
      return <ellipse cx={x + w / 2} cy={y + h / 2} rx={w / 2} ry={h / 2} {...paint} />
    case 'RoundRectangle':
      return <rect x={x} y={y} width={w} height={h} rx={Math.min(w, h) * 0.16} {...paint} />
    case 'Chevron':
      return <polygon points={`${x},${y} ${x + w - h / 2},${y} ${x + w},${y + h / 2} ${x + w - h / 2},${y + h} ${x},${y + h} ${x + h / 2},${y + h / 2}`} {...paint} />
    case 'Triangle':
      return <polygon points={`${x + w / 2},${y} ${x + w},${y + h} ${x},${y + h}`} {...paint} />
    default:
      return <rect x={x} y={y} width={w} height={h} {...paint} />
  }
}

function text(s: FakeShape): ReactNode {
  const frame = s.textFrame
  if (!frame) return null
  const content = frame.runs.map((r) => r.text).join('')
  if (!content.trim()) return null
  const justify = { Top: 'flex-start', Middle: 'center', Bottom: 'flex-end' }[frame.verticalAlignment] ?? 'flex-start'
  const align = frame.horizontalAlignment === 'Center' ? 'center' : frame.horizontalAlignment === 'Right' ? 'right'
    : /Justify|Distributed/.test(frame.horizontalAlignment) ? 'justify' : 'left'
  // Split runs into paragraphs at line breaks, keeping each run's font.
  const paragraphs: Array<Array<{ text: string; key: string; run: FakeRun }>> = [[]]
  frame.runs.forEach((r, i) => {
    r.text.split(/\r\n|\r|\n|\v/).forEach((piece, j) => {
      if (j > 0) paragraphs.push([])
      paragraphs[paragraphs.length - 1]?.push({ text: piece, key: `${i}-${j}`, run: r })
    })
  })
  return (
    <foreignObject x={s.left} y={s.top} width={s.width} height={Math.max(s.height, 1)} style={{ overflow: 'visible' }}>
      <div style={{
        width: '100%', height: '100%', boxSizing: 'border-box', display: 'flex', flexDirection: 'column', justifyContent: justify,
        padding: `${frame.topMargin}px ${frame.rightMargin}px ${frame.bottomMargin}px ${frame.leftMargin}px`, textAlign: align,
        overflowWrap: 'break-word', lineHeight: 1.18,
      }}>
        {paragraphs.map((p, i) => (
          <div key={i} style={{ display: 'flex', gap: 8, minHeight: '1em' }}>
            {frame.bulletVisible && p.some((x) => x.text) && <span style={{ color: '#328840', fontSize: 12 }}>■</span>}
            <span>
              {p.map(({ text: t, key, run }) => (
                <span key={key} style={{
                  fontFamily: `'${run.font.name ?? 'Aptos'}', Aptos, sans-serif`, fontSize: `${run.font.size ?? 18}px`,
                  color: run.font.color ?? '#000000', fontWeight: run.font.bold ? 700 : 400, fontStyle: run.font.italic ? 'italic' : 'normal',
                  whiteSpace: 'pre-wrap',
                }}>{t}</span>
              ))}
            </span>
          </div>
        ))}
      </div>
    </foreignObject>
  )
}

function table(s: FakeShape): ReactNode {
  const values = s.table?.values ?? []
  const rows = values.length
  const cols = values[0]?.length ?? 0
  if (!rows || !cols) return null
  const w = s.width / cols
  const h = s.height / rows
  return values.flatMap((row, r) =>
    row.map((cell, c) => (
      <g key={`${r}-${c}`}>
        <rect x={s.left + c * w} y={s.top + r * h} width={w} height={h} fill={r === 0 ? '#022338' : r % 2 ? '#FFFFFF' : '#F7F8F5'} stroke="#E8E9E6" />
        <text x={s.left + c * w + 8} y={s.top + r * h + h / 2 + 5} fontSize={13} fontFamily="Aptos, sans-serif" fill={r === 0 ? '#FFFFFF' : '#000000'} fontWeight={r === 0 ? 700 : 400}>{cell}</text>
      </g>
    )),
  )
}

function ShapeView({ model, shape: s, interactive, onPick }: { model: FakePresentation; shape: FakeShape; interactive: boolean; onPick: (e: MouseEvent, id: string) => void }) {
  if (!s.visible) return null
  const b = boundsOf(s)
  const transform = s.rotation ? `rotate(${s.rotation} ${b.left + b.width / 2} ${b.top + b.height / 2})` : undefined
  let body: ReactNode
  if (s.members?.length) body = s.members.map((m) => <ShapeView key={m.id} model={model} shape={m} interactive={false} onPick={onPick} />)
  else if (s.type === 'Line') body = <line x1={s.left} y1={s.top} x2={s.left + s.width} y2={s.top + s.height} stroke={s.line.color} strokeWidth={Math.max(1, s.line.weight)} />
  else if (s.type === 'Table') body = table(s)
  else if (s.type === 'Image' || s.type === 'Graphic') {
    const url = pictureUrl(model, s)
    body = url
      ? <image href={url} x={s.left} y={s.top} width={s.width} height={s.height} preserveAspectRatio="none" />
      : <rect x={s.left} y={s.top} width={s.width} height={s.height} fill="#E8E9E6" />
  } else body = <>{outline(s)}{text(s)}</>
  return (
    <g transform={transform} data-shape={s.id} onMouseDown={interactive ? (e) => onPick(e, s.id) : undefined} style={interactive ? { cursor: 'move' } : undefined}>
      {/* A clear frame, so a shape with no fill can still be clicked. */}
      {interactive && <rect x={b.left} y={b.top} width={Math.max(b.width, 4)} height={Math.max(b.height, 4)} fill="transparent" />}
      {body}
    </g>
  )
}

function Handles({ shape: s }: { shape: FakeShape }) {
  const b = boundsOf(s)
  const points = [[0, 0], [0.5, 0], [1, 0], [0, 0.5], [1, 0.5], [0, 1], [0.5, 1], [1, 1]] as const
  return (
    <g pointerEvents="none">
      <rect x={b.left} y={b.top} width={b.width} height={b.height} fill="none" stroke="#2B7CD3" strokeWidth={1.2} strokeDasharray="5 3" />
      {points.map(([fx, fy]) => <circle key={`${fx}-${fy}`} cx={b.left + b.width * fx} cy={b.top + b.height * fy} r={4} fill="#FFFFFF" stroke="#2B7CD3" strokeWidth={1.2} />)}
    </g>
  )
}

export function SlideView({ model, slide, interactive, onChange }: { model: FakePresentation; slide: FakeSlide; interactive: boolean; onChange?: () => void }) {
  const { slideWidth, slideHeight } = model.pageSetup
  const selected = new Set(interactive ? model.selection.shapeIds : [])
  const pick = (e: MouseEvent, id: string | null): void => {
    e.stopPropagation()
    const now = model.selection.shapeIds
    const add = e.shiftKey || e.ctrlKey || e.metaKey
    const next = id === null ? [] : add ? (now.includes(id) ? now.filter((x) => x !== id) : [...now, id]) : [id]
    void clickShapes(model, next, slide.id).then(onChange)
  }
  return (
    <svg viewBox={`0 0 ${slideWidth} ${slideHeight}`} className="demo-slide" onMouseDown={interactive ? (e) => pick(e, null) : undefined}>
      <rect width={slideWidth} height={slideHeight} fill="#FFFFFF" />
      {slide.shapes.map((s) => <ShapeView key={s.id} model={model} shape={s} interactive={interactive} onPick={pick} />)}
      {slide.shapes.filter((s) => selected.has(s.id)).map((s) => <Handles key={s.id} shape={s} />)}
    </svg>
  )
}

export function SlideStage({ model, onChange }: { model: FakePresentation; onChange: () => void }) {
  const active = activeSlide(model) ?? model.slides[0]
  const choose = (id: string): void => {
    selectSlides(model, [id])
    void fireSelectionChanged(model).then(onChange)
  }
  return (
    <div className="demo-ppt-stage">
      <ol className="demo-thumbs">
        {model.slides.map((s, i) => (
          <li key={s.id}>
            <button type="button" aria-current={s.id === active?.id} onClick={() => choose(s.id)} aria-label={`Slide ${i + 1}`}>
              <span>{i + 1}</span>
              <SlideView model={model} slide={s} interactive={false} />
            </button>
          </li>
        ))}
      </ol>
      <div className="demo-editor">
        {active && <SlideView model={model} slide={active} interactive onChange={onChange} />}
        <p className="demo-hint">
          A pretend PowerPoint for trying the pane. Click a shape to select it; Shift-click to add more. Then use the Arrange tab.
        </p>
      </div>
    </div>
  )
}
