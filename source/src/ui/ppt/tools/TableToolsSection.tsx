import { useState } from 'react'
import { hasPowerPoint, supportsPpt, updateMessage } from '../../../ppt/run.ts'
import {
  addTableTotalRow, equaliseColumnWidths, formatSelectedTable, pasteIntoSelectedTable, pasteNewTable, readTableHeaders, setColumnWidths,
  sortSelectedTable, transposeSelectedTable,
} from '../../../ppt/tools/tables.ts'
import { Button, Section, SelectField, TextField } from '../../components.tsx'
import { useAction } from '../../useAction.ts'
import './tools.css'

type Order = 'ascending' | 'descending'

const ORDERS: Array<{ value: Order; label: string }> = [
  { value: 'ascending', label: 'Smallest, or A, first' },
  { value: 'descending', label: 'Largest, or Z, first' },
]

/** Tools for the selected PowerPoint table, and new tables from cells copied from Excel. */
export function TableToolsSection() {
  const { busy, run } = useAction()
  const [headers, setHeaders] = useState<string[]>([])
  const [column, setColumn] = useState('0')
  const [order, setOrder] = useState<Order>('ascending')
  const [widths, setWidths] = useState('')
  const [pasted, setPasted] = useState('')
  const live = hasPowerPoint()
  const old18 = live && !supportsPpt('1.8')
  const old19 = live && !supportsPpt('1.9')
  const off = busy !== null || old18
  const cells = old19 ? updateMessage('Formatting cells, adding rows and setting column widths') : undefined
  const go = (key: string, action: () => Promise<string>) => (): void => void run(key, action)

  const columns = headers.length ? headers.map((h, i) => ({ value: String(i), label: h })) : [{ value: '0', label: 'First column' }]

  return (
    <Section id="tables" title="Tables" intro="Select a table by clicking its edge, then choose a tool.">
      {old18 && <p className="tools-note">{updateMessage('The table tools')}</p>}
      <div className="tool-row">
        <Button disabled={off || old19} title={cells ?? 'Navy header row, light grey rules, figures right-aligned'} onClick={go('format', formatSelectedTable)}>BwB format</Button>
        <Button disabled={off} title="Rows become columns" onClick={go('transpose', transposeSelectedTable)}>Transpose</Button>
        <Button disabled={off} title="Adds up each column of figures" onClick={go('total', addTableTotalRow)}>Add total row</Button>
        <Button disabled={off || old19} title={cells ?? 'Every column the same width'} onClick={go('equal', equaliseColumnWidths)}>Equal widths</Button>
      </div>
      {cells && <p className="tools-note">{cells}</p>}

      <h3>Sort</h3>
      <div className="field-pair">
        <SelectField label="By column" value={column} options={columns} onChange={setColumn} />
        <SelectField label="Order" value={order} options={ORDERS} onChange={setOrder} />
      </div>
      <div className="tool-row">
        <Button small disabled={off} onClick={go('headers', async () => {
          const found = await readTableHeaders()
          setHeaders(found)
          setColumn('0')
          return `${found.length} columns read from the table.`
        })}>Read the columns</Button>
        <Button small variant="navy" disabled={off} onClick={go('sort', () => sortSelectedTable(Number(column), order === 'descending'))}>Sort</Button>
      </div>

      <h3>Column widths</h3>
      <TextField label="Widths in centimetres, one for each column" value={widths} onChange={setWidths} placeholder="4, 3, 3, 3" />
      <Button small disabled={off || old19} title={cells} onClick={go('widths', () => setColumnWidths(widths))}>Set widths</Button>

      <h3>Paste from Excel</h3>
      <div className="paste-input">
        <TextField multiline label="Cells copied from Excel" value={pasted} onChange={setPasted} placeholder="Copy cells in Excel and paste them here" />
      </div>
      <div className="tool-row">
        <Button variant="navy" disabled={off || !pasted.trim()} onClick={go('paste-new', () => pasteNewTable(pasted))}>New BwB table</Button>
        <Button disabled={off || !pasted.trim()} onClick={go('paste-into', () => pasteIntoSelectedTable(pasted))}>{'Into the selected table'}</Button>
      </div>
    </Section>
  )
}
