import { useState } from 'react'
import { hasPowerPoint, supportsPpt, updateMessage } from '../../../ppt/run.ts'
import { MAX_PEOPLE, insertOrgChart } from '../../../ppt/tools/orgChart.ts'
import { Button, Section, TextField } from '../../components.tsx'
import { useAction } from '../../useAction.ts'
import './tools.css'

const EXAMPLE = [
  'Alex Morgan, Managing Director',
  '  Sam Taylor, Head of Advisory',
  '    Jo Patel, Associate',
  '  Chris Lee, Head of Operations',
].join('\n')

/** An org chart from a list typed or pasted into the pane. */
export function OrgChartSection() {
  const { busy, run } = useAction()
  const [text, setText] = useState('')
  const blocked = hasPowerPoint() && !supportsPpt('1.5')
  return (
    <Section id="org-chart" title="Org chart"
      intro={`One person a line, as Name, Role. Indent people under their manager, or write Name | Role | Reports to. Up to ${MAX_PEOPLE} people.`}>
      {blocked && <p className="tools-note">{updateMessage('The org chart')}</p>}
      <div className="org-input">
        <TextField multiline label="People" value={text} onChange={setText} placeholder={EXAMPLE} />
      </div>
      <Button variant="navy" disabled={busy !== null || blocked || !text.trim()} busy={busy === 'org'} onClick={() => void run('org', () => insertOrgChart(text))}>
        Add the org chart
      </Button>
    </Section>
  )
}
