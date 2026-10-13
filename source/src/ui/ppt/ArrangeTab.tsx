import { useState } from 'react'
import type { ReactNode } from 'react'
import {
  AlignHorizontalJustifyCenter, AlignHorizontalJustifyEnd, AlignHorizontalJustifyStart, AlignVerticalJustifyCenter, AlignVerticalJustifyEnd,
  AlignVerticalJustifyStart, ArrowDown, ArrowDownToLine, ArrowLeft, ArrowLeftRight, ArrowLeftToLine, ArrowRight, ArrowRightToLine, ArrowUp,
  ArrowUpToLine, BringToFront, ChevronDown, ChevronUp, Columns3, Grid2x2, Grid3x3, Group, LayoutGrid, Rows3, SendToBack, Ungroup,
} from 'lucide-react'
import type { LucideIcon } from 'lucide-react'
import { GUTTER } from '../../brand/slideGrid.ts'
import type { AlignEdge, Axis, Edge, Extent, Reference, SizeSource } from '../../geometry/index.ts'
import { cmToPt, ptToCm } from '../../geometry/index.ts'
import type { ZOrder } from '../../ppt/arrange.ts'
import {
  alignShapes, distributeShapes, equaliseShapeGaps, gridShapes, groupShapes, matchShapeSize, nudgeShapes, orderShapes, readShapeFrame,
  setShapeFrame, snapShapes, stackShapes, stretchShapes, swapShapes, tidyShapes, ungroupShapes,
} from '../../ppt/arrange.ts'
import type { VerticalAnchor } from '../../ppt/format.ts'
import { setAutofitOff, setTextMargins, setVerticalAnchor, setWordWrap } from '../../ppt/format.ts'
import { usePref, useSelection } from '../../ppt/hooks.ts'
import { hasPowerPoint, supportsPpt, updateMessage } from '../../ppt/run.ts'
import type { SelectionState } from '../../ppt/selection.ts'
import { selectionTracker } from '../../ppt/selection.ts'
import { UserFacingError } from '../../word/run.ts'
import { Button, Checkbox, Notice, Section, SelectField, TextField } from '../components.tsx'
import { useAction } from '../useAction.ts'
import { FormatPainterSection } from './tools/FormatPainterSection.tsx'
import { SelectSimilarSection } from './tools/SelectSimilarSection.tsx'
import './ppt.css'

type Option<T> = { value: T; label: string }

const REFERENCES: Array<Option<Reference>> = [
  { value: 'selection', label: 'Selection' }, { value: 'anchor', label: 'Anchor' }, { value: 'slide', label: 'Slide' },
  { value: 'margins', label: 'Margins' }, { value: 'content', label: 'Content' },
]
const AXES: Array<Option<Axis>> = [{ value: 'horizontal', label: 'Left to right' }, { value: 'vertical', label: 'Top to bottom' }]
const SPACINGS: Array<Option<'gaps' | 'centres'>> = [{ value: 'gaps', label: 'Equal gaps' }, { value: 'centres', label: 'Equal centres' }]
const EXTENTS: Array<Option<Extent>> = [
  { value: 'selection', label: 'Outer shapes' }, { value: 'margins', label: 'Margins' },
  { value: 'content', label: 'Content area' }, { value: 'slide', label: 'Whole slide' },
]
const SOURCES: Array<Option<SizeSource>> = [
  { value: 'anchor', label: 'The anchor' }, { value: 'largest', label: 'The largest' }, { value: 'smallest', label: 'The smallest' },
]
const ALIGN: Array<{ edge: AlignEdge; label: string; Icon: LucideIcon }> = [
  { edge: 'left', label: 'Left', Icon: AlignHorizontalJustifyStart }, { edge: 'centre', label: 'Centre', Icon: AlignHorizontalJustifyCenter },
  { edge: 'right', label: 'Right', Icon: AlignHorizontalJustifyEnd }, { edge: 'top', label: 'Top', Icon: AlignVerticalJustifyStart },
  { edge: 'middle', label: 'Middle', Icon: AlignVerticalJustifyCenter }, { edge: 'bottom', label: 'Bottom', Icon: AlignVerticalJustifyEnd },
]
const STRETCH: Array<{ edge: Edge; label: string; Icon: LucideIcon }> = [
  { edge: 'left', label: 'Left edge', Icon: ArrowLeftToLine }, { edge: 'right', label: 'Right edge', Icon: ArrowRightToLine },
  { edge: 'top', label: 'Top edge', Icon: ArrowUpToLine }, { edge: 'bottom', label: 'Bottom edge', Icon: ArrowDownToLine },
]
const ORDER: Array<{ z: ZOrder; label: string; Icon: LucideIcon }> = [
  { z: 'BringToFront', label: 'To front', Icon: BringToFront }, { z: 'BringForward', label: 'Forward', Icon: ChevronUp },
  { z: 'SendBackward', label: 'Backward', Icon: ChevronDown }, { z: 'SendToBack', label: 'To back', Icon: SendToBack },
]
const NUDGES: Array<{ key: string; label: string; Icon: LucideIcon; dx: number; dy: number }> = [
  { key: 'up', label: 'Up', Icon: ArrowUp, dx: 0, dy: -1 }, { key: 'left', label: 'Left', Icon: ArrowLeft, dx: -1, dy: 0 },
  { key: 'down', label: 'Down', Icon: ArrowDown, dx: 0, dy: 1 }, { key: 'right', label: 'Right', Icon: ArrowRight, dx: 1, dy: 0 },
]

/** Centimetres typed in the pane (a decimal comma is fine), or undefined for an empty field. */
function cm(text: string, what: string): number | undefined {
  const t = text.trim().replace(',', '.')
  if (!t) return undefined
  const n = Number(t)
  if (!Number.isFinite(n)) throw new UserFacingError(`${what}: “${text}” is not a number of centimetres.`)
  return n
}

const toCm = (pt: number): string => String(Math.round(ptToCm(pt) * 100) / 100)

function anchorLine(s: SelectionState): ReactNode {
  if (s.anchor && s.anchor.slideId === s.slideId) return <>Anchor: <b>{`“${s.anchor.name}”`}</b>, set by you</>
  const first = s.order[0]
  if (first) return <>Anchor: <b>{`“${s.names[first] ?? 'first shape'}”`}</b>, the first you selected</>
  return <>Anchor: the largest shape you select</>
}

function Segmented<T extends string>({ label, value, options, onChange }: { label: string; value: T; options: Array<Option<T>>; onChange: (v: T) => void }) {
  return (
    <div className="segmented" role="group" aria-label={label}>
      {options.map((o) => (
        <button key={o.value} type="button" aria-pressed={value === o.value} onClick={() => onChange(o.value)}>{o.label}</button>
      ))}
    </div>
  )
}

/** An icon button with its label under the icon, or the icon alone (the label then names it for screen readers). */
function Tool({ label, Icon, onClick, disabled, title, iconOnly, className }: {
  label: string
  Icon: LucideIcon
  onClick: () => void
  disabled: boolean
  title?: string
  iconOnly?: boolean
  className?: string
}) {
  return (
    <button type="button" className={`btn ${className ?? ''}`} onClick={onClick} disabled={disabled} title={title ?? label} aria-label={iconOnly ? label : undefined}>
      <Icon size={17} strokeWidth={1.8} aria-hidden />
      {iconOnly ? null : label}
    </button>
  )
}

/** The power alignment panel: what to line up with, then align, distribute, size, position, layout and order. */
export function ArrangeTab() {
  const { busy, run } = useAction()
  const selection = useSelection()
  const [reference, setReference] = usePref<Reference>('bwb.ppt.reference', 'selection', REFERENCES.map((r) => r.value))
  const [asGroup, setAsGroup] = usePref<boolean>('bwb.ppt.asGroup', false)
  const [axis, setAxis] = usePref<Axis>('bwb.ppt.axis', 'horizontal', AXES.map((a) => a.value))
  const [spacing, setSpacing] = usePref<'gaps' | 'centres'>('bwb.ppt.spacing', 'gaps', SPACINGS.map((s) => s.value))
  const [within, setWithin] = usePref<Extent>('bwb.ppt.within', 'selection', EXTENTS.map((e) => e.value))
  const [source, setSource] = usePref<SizeSource>('bwb.ppt.sizeSource', 'anchor', SOURCES.map((s) => s.value))
  const [keepAspect, setKeepAspect] = usePref<boolean>('bwb.ppt.keepAspect', false)
  const [uniform, setUniform] = usePref<boolean>('bwb.ppt.uniformCells', false)
  const [step, setStep] = usePref<string>('bwb.ppt.nudgeCm', '0.1')
  const [gap, setGap] = useState('')
  const [frame, setFrame] = useState({ x: '', y: '', w: '', h: '' })
  const [stackGap, setStackGap] = useState('0')
  const [columns, setColumns] = useState('3')
  const [gridGap, setGridGap] = useState(toCm(GUTTER))

  const live = hasPowerPoint()
  const ready = !live || supportsPpt('1.5')
  const can18 = !live || supportsPpt('1.8')
  const off = busy !== null || !ready
  const go = (key: string, action: () => Promise<string | void>) => (): void => void run(key, action)
  const count = selection.order.length
  const orderReason = can18 ? undefined : updateMessage('Changing the order of shapes or grouping them')

  const readFrame = async (): Promise<string> => {
    const f = await readShapeFrame()
    if (!f) return 'Select a shape first.'
    setFrame({ x: toCm(f.left), y: toCm(f.top), w: toCm(f.width), h: toCm(f.height) })
    return 'Position and size read from the selection.'
  }
  const applyFrame = async (): Promise<string> => {
    const values = { left: cm(frame.x, 'X'), top: cm(frame.y, 'Y'), width: cm(frame.w, 'Width'), height: cm(frame.h, 'Height') }
    const pts = Object.fromEntries(Object.entries(values).filter(([, v]) => v !== undefined).map(([k, v]) => [k, cmToPt(v ?? 0)]))
    if (!Object.keys(pts).length) throw new UserFacingError('Type a position or a size first.')
    return setShapeFrame(pts)
  }
  const spread = (): Promise<string> => {
    const exact = cm(gap, 'Exact gap')
    return distributeShapes({ axis, spacing, within, ...(exact === undefined ? {} : { gap: cmToPt(exact) }) })
  }
  const tile = (): Promise<string> => {
    const n = Number(columns)
    if (!Number.isInteger(n) || n < 1) throw new UserFacingError('Type a whole number of columns, 1 or more.')
    const g = cmToPt(cm(gridGap, 'Gap') ?? 0)
    return gridShapes({ columns: n, gapX: g, gapY: g, cells: uniform ? 'uniform' : 'keep' })
  }
  const nudgeBy = (dx: number, dy: number) => (): Promise<string> => {
    const s = cmToPt(cm(step, 'Step') ?? 0)
    return nudgeShapes(dx * s, dy * s)
  }

  return (
    <>
      {!ready && <Notice kind="warn">{updateMessage('Arranging shapes')}</Notice>}
      <div className="selection-bar" aria-live="polite">
        <strong>{count ? `${count} ${count === 1 ? 'shape' : 'shapes'} selected` : 'Nothing selected'}</strong>
        <span className="anchor">{anchorLine(selection)}</span>
        <div className="row">
          <Button small onClick={go('anchor', () => selectionTracker.setAnchorFromSelection())} disabled={off}>Set anchor</Button>
          <Button small variant="link" disabled={off || !selection.anchor} onClick={go('clear', async () => {
            selectionTracker.clearAnchor()
            return 'Anchor cleared. The first shape you select is the anchor again.'
          })}>Clear</Button>
        </div>
      </div>

      <SelectSimilarSection />

      <Section title="Line up" weight="act" intro="Choose what to line up with, then an edge. The anchor never moves.">
        <Segmented label="Line up with" value={reference} options={REFERENCES} onChange={setReference} />
        <div className="tool-grid">
          {ALIGN.map(({ edge, label, Icon }) => (
            <Tool key={edge} label={label} Icon={Icon} disabled={off} title={`Line up ${label.toLowerCase()} with the ${reference}`}
              onClick={go(`align-${edge}`, () => alignShapes(edge, reference, asGroup))} />
          ))}
        </div>
        <Checkbox label="Move the shapes as a group" checked={asGroup} onChange={setAsGroup} hint="Keeps their layout, as for centring a whole diagram on the slide." />
        {live && !supportsPpt('1.10') && <p className="needs-update">This PowerPoint does not report rotation, so turned shapes are lined up as if upright.</p>}
      </Section>

      <Section title="Spread out">
        <Segmented label="Direction" value={axis} options={AXES} onChange={setAxis} />
        <Segmented label="Spacing" value={spacing} options={SPACINGS} onChange={setSpacing} />
        <div className="field-pair">
          <SelectField label="Within" value={within} options={EXTENTS} onChange={setWithin} />
          <TextField label="Exact gap (cm), optional" value={gap} onChange={setGap} placeholder="Leave empty" />
        </div>
        <div className="row">
          <Button variant="navy" disabled={off} onClick={go('distribute', spread)}>Spread out</Button>
          <Button disabled={off} onClick={go('equalise', () => equaliseShapeGaps(axis))} title="Every gap becomes the average gap; the first shape stays">
            Even out gaps
          </Button>
        </div>
      </Section>

      <Section title="Same size">
        <div className="field-pair">
          <SelectField label="Size from" value={source} options={SOURCES} onChange={setSource} />
          <Checkbox label="Keep proportions" checked={keepAspect} onChange={setKeepAspect} />
        </div>
        <div className="row">
          <Button disabled={off} onClick={go('width', () => matchShapeSize('width', source, keepAspect))}>Same width</Button>
          <Button disabled={off} onClick={go('height', () => matchShapeSize('height', source, keepAspect))}>Same height</Button>
          <Button disabled={off} onClick={go('size', () => matchShapeSize('both', source))}>Same size</Button>
        </div>
      </Section>

      <Section title="Position and size" intro="In centimetres from the top-left corner of the slide. Empty fields stay as they are.">
        <div className="cm-fields">
          <TextField label="X" value={frame.x} onChange={(x) => setFrame({ ...frame, x })} />
          <TextField label="Y" value={frame.y} onChange={(y) => setFrame({ ...frame, y })} />
          <TextField label="Width" value={frame.w} onChange={(w) => setFrame({ ...frame, w })} />
          <TextField label="Height" value={frame.h} onChange={(h) => setFrame({ ...frame, h })} />
        </div>
        <div className="row">
          <Button disabled={off} onClick={go('read', readFrame)}>Read from selection</Button>
          <Button variant="navy" disabled={off} onClick={go('apply', applyFrame)}>Apply</Button>
        </div>
      </Section>

      <Section title="Layout">
        <div className="field-pair">
          <TextField label="Stack gap (cm)" value={stackGap} onChange={setStackGap} />
          <div className="row">
            <Tool label="Side by side" Icon={Columns3} disabled={off} onClick={go('stack-h', () => stackShapes('horizontal', cmToPt(cm(stackGap, 'Stack gap') ?? 0)))} />
            <Tool label="Top to bottom" Icon={Rows3} disabled={off} onClick={go('stack-v', () => stackShapes('vertical', cmToPt(cm(stackGap, 'Stack gap') ?? 0)))} />
          </div>
        </div>
        <div className="field-pair">
          <TextField label="Grid columns" value={columns} onChange={setColumns} type="number" />
          <TextField label="Grid gap (cm)" value={gridGap} onChange={setGridGap} />
        </div>
        <Checkbox label="Same size cells" checked={uniform} onChange={setUniform} hint={'Every shape takes the first shape\u2019s size.'} />
        <div className="tool-grid four">
          <Tool label="Grid" Icon={Grid2x2} disabled={off} onClick={go('grid', tile)} title="Arrange in a grid, in reading order" />
          <Tool label="Tidy up" Icon={LayoutGrid} disabled={off} onClick={go('tidy', tidyShapes)} title="Find the rows and columns and line them up with equal gaps" />
          <Tool label="Snap" Icon={Grid3x3} disabled={off} onClick={go('snap', snapShapes)} title="Snap the edges to the BwB 12-column grid" />
          <Tool label="Swap" Icon={ArrowLeftRight} disabled={off} onClick={go('swap', swapShapes)} title="Swap the places of two shapes" />
        </div>
      </Section>

      <Section title="Stretch" intro={`Moves one edge of each shape to the same edge of the ${reference}, keeping the opposite edge.`}>
        <div className="tool-grid four">
          {STRETCH.map(({ edge, label, Icon }) => (
            <Tool key={edge} label={label} Icon={Icon} disabled={off} onClick={go(`stretch-${edge}`, () => stretchShapes(edge, reference))} />
          ))}
        </div>
      </Section>

      <Section title="Nudge">
        <div className="field-pair">
          <TextField label="Step (cm)" value={step} onChange={setStep} />
          <div className="nudge-pad">
            {NUDGES.map(({ key, label, Icon, dx, dy }) => (
              <Tool key={key} className={key} iconOnly label={`Nudge ${label.toLowerCase()}`} Icon={Icon} disabled={off} onClick={go(`nudge-${key}`, nudgeBy(dx, dy))} />
            ))}
          </div>
        </div>
      </Section>

      <Section title="Order and group">
        {orderReason && <p className="needs-update">{orderReason}</p>}
        <div className="tool-grid">
          {ORDER.map(({ z, label, Icon }) => (
            <Tool key={z} label={label} Icon={Icon} disabled={off || !can18} title={orderReason ?? label} onClick={go(z, () => orderShapes(z))} />
          ))}
          <Tool label="Group" Icon={Group} disabled={off || !can18} title={orderReason ?? 'Group the selected shapes'} onClick={go('group', groupShapes)} />
          <Tool label="Ungroup" Icon={Ungroup} disabled={off || !can18} title={orderReason ?? 'Ungroup the selected groups'} onClick={go('ungroup', ungroupShapes)} />
        </div>
      </Section>

      <FormatPainterSection />

      <Section title="Text boxes" intro="For the selected shapes and text boxes.">
        <div className="row">
          <Button small disabled={off} onClick={go('margins-zero', () => setTextMargins('zero'))}>No margins</Button>
          <Button small disabled={off} onClick={go('margins-std', () => setTextMargins('standard'))}>Margins 0.2 cm</Button>
          <Button small disabled={off} onClick={go('autofit', setAutofitOff)}>Autofit off</Button>
          <Button small disabled={off} onClick={go('wrap-on', () => setWordWrap(true))}>Wrap text</Button>
          <Button small disabled={off} onClick={go('wrap-off', () => setWordWrap(false))}>Do not wrap</Button>
        </div>
        <h3>Anchor the text</h3>
        <div className="row">
          {(['Top', 'Middle', 'Bottom'] as VerticalAnchor[]).map((v) => (
            <Button key={v} small disabled={off} onClick={go(`anchor-${v}`, () => setVerticalAnchor(v))}>{v}</Button>
          ))}
        </div>
      </Section>
    </>
  )
}
