// The editable data grid behind the chart builder: rows of cell text, the first row for series names and the first
// column for categories. Every operation returns a new grid, so React state stays immutable.

export type Grid = string[][]

/** A rectangular copy with at least one cell. */
export function normaliseGrid(grid: Grid): Grid {
  const width = Math.max(1, ...grid.map((row) => row.length))
  const rows = grid.length ? grid : [['']]
  return rows.map((row) => Array.from({ length: width }, (_, c) => row[c] ?? ''))
}

const isBlank = (cells: string[]): boolean => cells.every((c) => !c.trim())

/** Drops empty rows and columns from the end, keeping at least one cell. */
export function trimGrid(grid: Grid): Grid {
  let rows = normaliseGrid(grid)
  while (rows.length > 1 && isBlank(rows[rows.length - 1] ?? [])) rows = rows.slice(0, -1)
  let width = rows[0]?.length ?? 1
  while (width > 1 && isBlank(rows.map((r) => r[width - 1] ?? ''))) width -= 1
  return rows.map((r) => r.slice(0, width))
}

export const cellAt = (grid: Grid, row: number, column: number): string => grid[row]?.[column] ?? ''

export function setCell(grid: Grid, row: number, column: number, value: string): Grid {
  const out = normaliseGrid(grid).map((r) => [...r])
  while (out.length <= row) out.push(Array.from({ length: out[0]?.length ?? 1 }, () => ''))
  const width = Math.max(out[0]?.length ?? 1, column + 1)
  const sized = out.map((r) => Array.from({ length: width }, (_, c) => r[c] ?? ''))
  const target = sized[row]
  if (target) target[column] = value
  return sized
}

export function addRow(grid: Grid, at?: number): Grid {
  const rows = normaliseGrid(grid)
  const index = Math.min(rows.length, Math.max(0, at ?? rows.length))
  const blank = Array.from({ length: rows[0]?.length ?? 1 }, () => '')
  return [...rows.slice(0, index), blank, ...rows.slice(index)]
}

export function removeRow(grid: Grid, index: number): Grid {
  const rows = normaliseGrid(grid)
  if (rows.length <= 1 || index < 0 || index >= rows.length) return rows
  return rows.filter((_, r) => r !== index)
}

export function addColumn(grid: Grid, at?: number): Grid {
  const rows = normaliseGrid(grid)
  const width = rows[0]?.length ?? 1
  const index = Math.min(width, Math.max(0, at ?? width))
  return rows.map((r) => [...r.slice(0, index), '', ...r.slice(index)])
}

export function removeColumn(grid: Grid, index: number): Grid {
  const rows = normaliseGrid(grid)
  const width = rows[0]?.length ?? 1
  if (width <= 1 || index < 0 || index >= width) return rows
  return rows.map((r) => r.filter((_, c) => c !== index))
}

/** Moves a column (a series) left or right, for the series order. */
export function moveColumn(grid: Grid, from: number, to: number): Grid {
  const rows = normaliseGrid(grid)
  const width = rows[0]?.length ?? 1
  if (from === to || from < 0 || to < 0 || from >= width || to >= width) return rows
  return rows.map((r) => {
    const copy = [...r]
    const [moved] = copy.splice(from, 1)
    copy.splice(to, 0, moved ?? '')
    return copy
  })
}

/** Swaps rows and columns: series become categories and the other way round. */
export function transpose(grid: Grid): Grid {
  const rows = normaliseGrid(grid)
  const width = rows[0]?.length ?? 1
  return Array.from({ length: width }, (_, c) => rows.map((r) => r[c] ?? ''))
}

/** Writes a pasted block into the grid with its top-left cell at row, column, growing the grid as needed. */
export function pasteBlock(grid: Grid, row: number, column: number, block: Grid): Grid {
  let out = normaliseGrid(grid)
  block.forEach((cells, r) => cells.forEach((value, c) => (out = setCell(out, row + r, column + c, value))))
  return out
}

const quoteCell = (cell: string): string => (/[\t\n"]/.test(cell) ? `"${cell.replace(/"/g, '""')}"` : cell)

/** Tab-separated text, the way Excel copies cells, so it can be pasted back into Excel. */
export function gridToText(grid: Grid): string {
  return trimGrid(grid).map((r) => r.map(quoteCell).join('\t')).join('\n')
}
