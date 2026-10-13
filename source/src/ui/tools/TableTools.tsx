import { useState } from 'react'
import { TEXT_WIDTH_CM } from '../../word/tools/common.ts'
import { readTableColumns, sortTable, tableToText, textToTable, transposeTable } from '../../word/tools/tableData.ts'
import {
  addTotalRow, centreTable, deleteEmptyRows, distributeColumns, fitTableToContents, fitTableToWindow, parseWidths, setColumnWidths,
} from '../../word/tools/tableTools.ts'
import { UserFacingError } from '../../word/run.ts'
import { unavailable } from '../../word/tools/support.ts'
import { Disclosure, Section, SelectField, TextField } from '../components.tsx'
import { useAction } from '../useAction.ts'
import { ToolButton } from './ToolButton.tsx'

type SortAs = 'text' | 'number'
type Order = 'ascending' | 'descending'

const DEFAULT_COLUMNS = ['Column 1', 'Column 2', 'Column 3', 'Column 4', 'Column 5', 'Column 6']

export function TableTools() {
  const { busy, run } = useAction()
  const working = busy !== null
  const reason = unavailable('core')
  const [widths, setWidths] = useState('')
  const [columns, setColumns] = useState<string[]>(DEFAULT_COLUMNS)
  const [column, setColumn] = useState('0')
  const [sortAs, setSortAs] = useState<SortAs>('text')
  const [order, setOrder] = useState<Order>('ascending')

  const tool = (key: string, label: string, action: () => Promise<string>, options: { variant?: 'primary' | 'navy'; reason?: string | null } = {}) => (
    <ToolButton key={key} label={label} busy={busy === key} working={working} reason={options.reason ?? reason}
      onClick={() => void run(key, action)} {...(options.variant ? { variant: options.variant } : {})} />
  )

  const readColumns = async (): Promise<string> => {
    const names = await readTableColumns()
    setColumns(names)
    setColumn('0')
    return `${names.length} columns read from the table at the cursor.`
  }

  const orderLabels = sortAs === 'number'
    ? { ascending: 'Smallest first', descending: 'Largest first' }
    : { ascending: 'A to Z', descending: 'Z to A' }

  return (
    <Section id="tables" title="Tables" intro="Put the cursor in a table first. Text to table works on the selected lines.">
      <div className="tool-row">
        {tool('distribute', 'Distribute columns', distributeColumns)}
        {tool('window', 'Fit to page width', fitTableToWindow)}
        {tool('centre', 'Centre', centreTable)}
      </div>
      <div className="tool-row">
        {tool('contents', 'Fit columns to contents', fitTableToContents, { reason: reason ?? unavailable('fitToContents') })}
      </div>
      <div className="tool-row">
        {tool('total', 'Add a total row', addTotalRow, { variant: 'primary' })}
        {tool('empty-rows', 'Delete empty rows', deleteEmptyRows)}
      </div>
      <Disclosure title="Column widths">
        <TextField label="Widths in cm" value={widths} onChange={setWidths} placeholder="4, 3, 3, 6.6"
          hint={`One width per column, first column first. The BwB text width is ${TEXT_WIDTH_CM} cm.`} />
        {tool('widths', 'Set column widths', async () => {
          const parsed = parseWidths(widths)
          if (!parsed) throw new UserFacingError('Give each width in centimetres, between 0.5 and 30, separated by commas: 4, 3, 3, 6.6')
          return setColumnWidths(parsed)
        }, { variant: 'navy' })}
      </Disclosure>
      <Disclosure title="Sort rows">
        <div className="tool-row">{tool('columns', 'Read the column headings', readColumns)}</div>
        <SelectField label="Sort by" value={column} onChange={setColumn}
          options={columns.map((name, i) => ({ value: String(i), label: name }))} />
        <div className="cols-2">
          <SelectField<SortAs> label="Sort as" value={sortAs} onChange={setSortAs}
            options={[{ value: 'text', label: 'Text' }, { value: 'number', label: 'Numbers' }]} />
          <SelectField<Order> label="Order" value={order} onChange={setOrder}
            options={[{ value: 'ascending', label: orderLabels.ascending }, { value: 'descending', label: orderLabels.descending }]} />
        </div>
        <p className="field-hint">The header row and a total row stay where they are. Cell shading stays in place.</p>
        {tool('sort', 'Sort rows', () => sortTable({ column: Number(column), descending: order === 'descending', numeric: sortAs === 'number' }), { variant: 'navy' })}
      </Disclosure>
      <Disclosure title="Convert">
        <div className="tool-row">{tool('to-table', 'Text to BwB table', textToTable, { variant: 'primary' })}</div>
        <p className="field-hint">Select lines with the columns separated by tabs, semicolons or commas. The first line becomes the header row.</p>
        <div className="tool-row">
          {tool('to-text', 'Table to text', tableToText)}
          {tool('transpose', 'Swap rows and columns', transposeTable)}
        </div>
      </Disclosure>
    </Section>
  )
}
