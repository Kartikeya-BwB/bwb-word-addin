import { useState } from 'react'
import { hasPowerPoint, supportsPpt, updateMessage } from '../../../ppt/run.ts'
import type { DiagramKind } from '../../../ppt/tools/diagrams.ts'
import { DIAGRAMS, insertDiagram } from '../../../ppt/tools/diagrams.ts'
import { Button, Section, SelectField } from '../../components.tsx'
import { useAction } from '../../useAction.ts'
import './tools.css'

/** Pyramid, funnel, cycle, Venn, and hub and spoke, with the number of items chosen in the pane. */
export function DiagramSection() {
  const { busy, run } = useAction()
  const [kind, setKind] = useState<DiagramKind>('pyramid')
  const info = DIAGRAMS.find((d) => d.kind === kind) ?? DIAGRAMS[0]
  const [count, setCount] = useState(String(info?.initial ?? 4))
  const live = hasPowerPoint()
  const blocked = live && !supportsPpt('1.5')
  const turns = !live || supportsPpt('1.10')
  if (!info) return null

  const choose = (next: DiagramKind): void => {
    setKind(next)
    setCount(String(DIAGRAMS.find((d) => d.kind === next)?.initial ?? 4))
  }
  const counts = Array.from({ length: info.max - info.min + 1 }, (_, i) => String(info.min + i)).map((n) => ({ value: n, label: n }))

  return (
    <Section id="diagrams" title="Diagrams" intro="Native shapes in BwB colours, placed in the content area, with text to type over.">
      {blocked && <p className="tools-note">{updateMessage('Diagrams')}</p>}
      <div className="field-pair">
        <SelectField label="Diagram" value={kind} options={DIAGRAMS.map((d) => ({ value: d.kind, label: d.label }))} onChange={choose} />
        <SelectField label={`Number of ${info.noun}`} value={count} options={counts} onChange={setCount} />
      </div>
      <p className="field-hint">{info.hint}.</p>
      {!turns && (kind === 'cycle' || kind === 'hub-spoke') && <p className="tools-note">This PowerPoint cannot turn shapes, so the arrows and slanted lines are left out.</p>}
      <Button variant="navy" disabled={busy !== null || blocked} busy={busy === 'diagram'} onClick={() => void run('diagram', () => insertDiagram(kind, Number(count)))}>
        {`Add the ${info.label.toLowerCase()}`}
      </Button>
    </Section>
  )
}
