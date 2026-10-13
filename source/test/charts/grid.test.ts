import { describe, expect, it } from 'vitest'
import {
  addColumn, addRow, gridToText, moveColumn, normaliseGrid, pasteBlock, removeColumn, removeRow, setCell, transpose, trimGrid,
} from '../../src/charts/grid.ts'
import { parsePasted } from '../../src/charts/data.ts'

const grid = [['Year', 'A', 'B'], ['2024', '1', '2'], ['2025', '3', '4']]

describe('grid model', () => {
  it('adds and removes rows', () => {
    expect(addRow(grid)).toEqual([...grid, ['', '', '']])
    expect(addRow(grid, 1)[1]).toEqual(['', '', ''])
    expect(removeRow(grid, 1)).toEqual([grid[0], grid[2]])
    expect(removeRow([['only']], 0)).toEqual([['only']])
  })

  it('adds and removes columns', () => {
    expect(addColumn(grid).map((r) => r.length)).toEqual([4, 4, 4])
    expect(addColumn(grid, 1)[0]).toEqual(['Year', '', 'A', 'B'])
    expect(removeColumn(grid, 1)).toEqual([['Year', 'B'], ['2024', '2'], ['2025', '4']])
    expect(removeColumn([['a'], ['b']], 0)).toEqual([['a'], ['b']])
  })

  it('transposes, and transposing twice gives the grid back', () => {
    expect(transpose(grid)).toEqual([['Year', '2024', '2025'], ['A', '1', '3'], ['B', '2', '4']])
    expect(transpose(transpose(grid))).toEqual(grid)
  })

  it('moves a column for the series order', () => {
    expect(moveColumn(grid, 2, 1)[0]).toEqual(['Year', 'B', 'A'])
    expect(moveColumn(grid, 2, 9)).toEqual(grid)
  })

  it('sets cells beyond the edge by growing the grid', () => {
    const grown = setCell(grid, 3, 4, 'x')
    expect(grown).toHaveLength(4)
    expect(grown.every((r) => r.length === 5)).toBe(true)
    expect(grown[3]?.[4]).toBe('x')
    expect(grid[0]).toEqual(['Year', 'A', 'B'])
  })

  it('pastes a block of cells at a position', () => {
    const pasted = pasteBlock(grid, 2, 1, [['9', '8', '7'], ['6', '5', '4']])
    expect(pasted).toEqual([
      ['Year', 'A', 'B', ''],
      ['2024', '1', '2', ''],
      ['2025', '9', '8', '7'],
      ['', '6', '5', '4'],
    ])
  })

  it('normalises ragged rows and trims empty edges', () => {
    expect(normaliseGrid([['a'], ['b', 'c']])).toEqual([['a', ''], ['b', 'c']])
    expect(normaliseGrid([])).toEqual([['']])
    expect(trimGrid([['a', '', ''], ['', '', ''], ['', '', '']])).toEqual([['a']])
  })

  it('writes tab-separated text that parses back to the same grid', () => {
    const tricky = [['Name', 'Note'], ['Tab\there', 'Say "hi"'], ['Two\nlines', '5']]
    const text = gridToText(tricky)
    expect(text.split('\n')[0]).toBe('Name\tNote')
    expect(parsePasted(text).grid).toEqual(tricky)
  })
})
