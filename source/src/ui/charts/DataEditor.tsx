// The chart's data: a box to paste cells from Excel into, and a compact grid to correct them in. Pasting a block of
// cells into a grid cell fills the grid from there, as in Excel.

import type { ClipboardEvent } from 'react'
import { useId } from 'react'
import type { Grid } from '../../charts/grid.ts'
import { addColumn, addRow, normaliseGrid, pasteBlock, removeColumn, removeRow, setCell, transpose } from '../../charts/grid.ts'
import { parsePasted } from '../../charts/data.ts'
import type { ChartType } from '../../charts/spec.ts'
import { Button } from '../components.tsx'

const MAX_ROWS = 60
const MAX_COLUMNS = 12

/** What each type expects, in the order people lay it out in Excel. */
export const LAYOUT_HINTS: Record<ChartType, string> = {
  column: 'Categories down the first column, one column of values per series, headings in the first row.',
  stackedColumn: 'Categories down the first column, one column per series to stack, headings in the first row.',
  column100: 'Categories down the first column, one column per series; each category is shown as shares of 100%.',
  bar: 'Categories down the first column, one column of values per series, headings in the first row.',
  stackedBar: 'Categories down the first column, one column per series to stack, headings in the first row.',
  bar100: 'Categories down the first column, one column per series; each category is shown as shares of 100%.',
  line: 'Periods down the first column, one column per line, headings in the first row. Leave a cell empty for a gap.',
  waterfall: 'Steps down the first column, values in the second, and \u{201C}total\u{201D} or \u{201C}subtotal\u{201D} in the third where needed. A total left blank is worked out.',
  mekko: 'Categories (the columns) down the first column, segments across the first row. Widths follow each category\u{2019}s total.',
  gantt: 'Tasks down the first column with start and end dates. One date makes a milestone; no dates makes a heading for the rows below.',
  donut: 'Segment names down the first column and their values in the second. Six segments at most.',
  harvey: 'Options down the first column, criteria across the first row, and a score from 0 to 4 in each cell.',
}

const columnName = (c: number): string => (c < 26 ? String.fromCharCode(65 + c) : `${c + 1}`)

export function DataEditor({ type, grid, text, onText, onGrid, onExample }: {
  type: ChartType
  grid: Grid
  text: string
  onText: (text: string) => void
  onGrid: (grid: Grid) => void
  onExample: () => void
}) {
  const pasteId = useId()
  const rows = normaliseGrid(grid)
  const shown = rows.slice(0, MAX_ROWS).map((r) => r.slice(0, MAX_COLUMNS))
  const width = shown[0]?.length ?? 1

  const onCellPaste = (event: ClipboardEvent<HTMLInputElement>, r: number, c: number): void => {
    const pasted = event.clipboardData.getData('text/plain')
    if (!/[\t\n]/.test(pasted.replace(/\r?\n$/, ''))) return
    event.preventDefault()
    onGrid(pasteBlock(rows, r, c, parsePasted(pasted).grid))
  }

  return (
    <div className="chart-data">
      <div className="field">
        <label htmlFor={pasteId}>Paste from Excel</label>
        <textarea id={pasteId} className="chart-paste" value={text} spellCheck={false} rows={4}
          placeholder="Copy the cells in Excel, headings included, and paste them here."
          onChange={(e) => onText(e.target.value)} />
        <span className="field-hint">{LAYOUT_HINTS[type]}</span>
      </div>
      <div className="chart-grid-wrap" role="region" aria-label="Chart data" tabIndex={0}>
        <table className="chart-grid">
          <thead>
            <tr>
              {Array.from({ length: width }, (_, c) => (
                <th key={c} scope="col">
                  {c > 0 && width > 2 ? (
                    <button type="button" className="chart-grid-remove" title={`Remove column ${columnName(c)}`}
                      aria-label={`Remove column ${columnName(c)}`} onClick={() => onGrid(removeColumn(rows, c))}>{'\u{00D7}'}</button>
                  ) : <span aria-hidden>{columnName(c)}</span>}
                </th>
              ))}
              <th aria-hidden />
            </tr>
          </thead>
          <tbody>
            {shown.map((cells, r) => (
              <tr key={r} className={r === 0 ? 'chart-grid-head' : undefined}>
                {cells.map((value, c) => (
                  <td key={c}>
                    <input value={value} aria-label={`Row ${r + 1}, column ${columnName(c)}`} spellCheck={false}
                      className={c === 0 || r === 0 ? 'chart-grid-label' : undefined}
                      onChange={(e) => onGrid(setCell(rows, r, c, e.target.value))}
                      onPaste={(e) => onCellPaste(e, r, c)} />
                  </td>
                ))}
                <td>
                  {r > 0 && rows.length > 2 && (
                    <button type="button" className="chart-grid-remove" title={`Remove row ${r + 1}`} aria-label={`Remove row ${r + 1}`}
                      onClick={() => onGrid(removeRow(rows, r))}>{'\u{00D7}'}</button>
                  )}
                </td>
              </tr>
            ))}
          </tbody>
        </table>
      </div>
      {(rows.length > MAX_ROWS || (rows[0]?.length ?? 0) > MAX_COLUMNS) && (
        <p className="field-hint">The grid shows the first {MAX_ROWS} rows and {MAX_COLUMNS} columns; the chart uses them all.</p>
      )}
      <div className="row chart-data-actions">
        <Button small onClick={() => onGrid(addRow(rows))}>Add row</Button>
        <Button small onClick={() => onGrid(addColumn(rows))}>Add column</Button>
        <Button small onClick={() => onGrid(transpose(rows))} title="Swap rows and columns">Transpose</Button>
        <Button small variant="link" onClick={onExample}>Example data</Button>
      </div>
    </div>
  )
}
