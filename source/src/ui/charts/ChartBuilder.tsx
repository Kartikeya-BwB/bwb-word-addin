// The chart builder: pick a type, paste or type the data, set the options, see the chart as it will look, and insert
// it. PowerPoint gets a vector graphic in the content area; Word gets a sharp picture with a numbered caption.

import './charts.css'
import { useMemo, useState } from 'react'
import { toSeriesTable } from '../../charts/data.ts'
import { isoDay, londonToday } from '../../charts/dates.ts'
import type { ChartType, RenderedChart } from '../../charts/spec.ts'
import { CHART_TYPES, renderChart } from '../../charts/spec.ts'
import { harveyTable } from '../../charts/types/harvey.ts'
import { currentHost } from '../../host.ts'
import { errorMessage } from '../../word/run.ts'
import { Button, Checkbox, Disclosure, Notice, SelectField, TextField } from '../components.tsx'
import { useAction } from '../useAction.ts'
import type { BuilderState, SizePreset } from './builderState.ts'
import { initialState, placement, sizeOptions, textOf, toSpec, withExample, withGrid, withPaste, withType } from './builderState.ts'
import { ChartOptions } from './ChartOptions.tsx'
import { DataEditor } from './DataEditor.tsx'
import { insertChart } from './insertChart.ts'
import type { RecentChart } from './recent.ts'
import { loadRecent, recentLabel, saveRecent } from './recent.ts'
import { TypeIcon } from './TypeIcons.tsx'

const dataUri = (svg: string): string => `data:image/svg+xml;charset=utf-8,${encodeURIComponent(svg)}`

const cm = (pt: number): string => (Math.round((pt * 2.54 * 10) / 72) / 10).toLocaleString('en-GB')

function draw(state: BuilderState, host: ReturnType<typeof currentHost>, today: string): { chart: RenderedChart | null; error: string | null } {
  try {
    return { chart: renderChart(toSpec(state, host, today)), error: null }
  } catch (error) {
    return { chart: null, error: errorMessage(error) }
  }
}

function savedOn(iso: string): string {
  const date = new Date(iso)
  return Number.isNaN(date.getTime()) ? '' : date.toLocaleDateString('en-GB', { day: 'numeric', month: 'long', year: 'numeric' })
}

export function ChartBuilder() {
  const { busy, run } = useAction()
  const [host] = useState(currentHost)
  const [state, setState] = useState<BuilderState>(() => initialState(host))
  const [text, setText] = useState(() => textOf(initialState(host)))
  const [recent, setRecent] = useState<RecentChart[]>(() => loadRecent(host))
  const today = isoDay(londonToday())
  const { chart, error } = useMemo(() => draw(state, host, today), [state, host, today])
  const table = useMemo(() => toSeriesTable(state.grid), [state.grid])
  const categories = state.type === 'harvey' ? harveyTable(state.grid).options : table.categories
  const place = placement(state, host)
  const word = host === 'word'

  const replace = (next: BuilderState): void => {
    setState(next)
    setText(textOf(next))
  }
  const patch = (changes: Partial<BuilderState>): void => setState({ ...state, ...changes })
  const pickType = (type: ChartType): void => {
    const next = withType(state, type)
    if (next.example) replace(next)
    else setState(next)
  }

  const insert = (): void =>
    void run('chart', async () => {
      if (!chart) throw new Error(error ?? 'The chart could not be drawn.')
      const message = await insertChart(chart, {
        host, title: state.title, source: state.source, sourceInChart: state.sourceInChart,
        ...(place.left !== undefined ? { left: place.left } : {}),
        ...(place.top !== undefined ? { top: place.top } : {}),
      })
      setRecent(saveRecent(state, host))
      return message
    })

  return (
    <div className="chart-builder">
      <div className="chart-types" role="group" aria-label="Chart type">
        {CHART_TYPES.map((t) => (
          <button key={t.type} type="button" aria-pressed={state.type === t.type} className="chart-type" title={t.hint}
            onClick={() => pickType(t.type)}>
            <TypeIcon type={t.type} />
            <span>{t.label}</span>
          </button>
        ))}
      </div>

      <figure className="chart-preview">
        {chart ? <img src={dataUri(chart.svg)} alt={chart.altText} /> : <Notice kind="error">{error}</Notice>}
        <figcaption>
          {word ? `${cm(place.widthPt)} \u{00D7} ${cm(place.heightPt)} cm` : `${place.widthPt} \u{00D7} ${place.heightPt} pt`}, labels {place.fontPt} pt
        </figcaption>
      </figure>
      {chart && chart.warnings.length > 0 && (
        <Notice kind="warn">{chart.warnings.map((w) => <span key={w} className="chart-warning">{w}</span>)}</Notice>
      )}

      <TextField label="Title" value={state.title} onChange={(title) => patch({ title })}
        placeholder="Climate finance by region, 2025"
        hint={word ? 'Goes in the caption below the chart as \u{201C}Figure 1. Title\u{201D}; it is never drawn on the chart.' : 'Used for the alt text. The slide title carries the message.'} />
      <TextField label="Source" value={state.source} onChange={(source) => patch({ source })} placeholder="OECD (2026)" />
      <Checkbox label="Source line inside the chart" checked={state.sourceInChart} onChange={(sourceInChart) => patch({ sourceInChart })}
        hint={word ? 'Otherwise it goes in the caption.' : '8 pt charcoal at the bottom left.'} />

      <h3>Data</h3>
      <DataEditor type={state.type} grid={state.grid} text={text}
        onText={(value) => {
          setText(value)
          setState(withPaste(state, value))
        }}
        onGrid={(grid) => {
          setState(withGrid(state, grid))
          setText(textOf({ ...state, grid }))
        }}
        onExample={() => replace(withExample(state))} />

      <Disclosure title="Options" open>
        <ChartOptions state={state} onChange={patch} categories={categories} seriesNames={table.series.map((s) => s.name)} />
      </Disclosure>

      <div className="chart-insert">
        <SelectField<SizePreset> label="Size" value={state.size} options={sizeOptions(host)} onChange={(size) => patch({ size })} />
        {state.size === 'custom' && (
          <div className="cols-2">
            <TextField label={word ? 'Width (cm)' : 'Width (pt)'} value={state.customWidth} onChange={(customWidth) => patch({ customWidth })} />
            <TextField label={word ? 'Height (cm)' : 'Height (pt)'} value={state.customHeight} onChange={(customHeight) => patch({ customHeight })} />
          </div>
        )}
        <Button variant="primary" block busy={busy === 'chart'} disabled={!chart} onClick={insert}>
          {word && state.title.trim() ? 'Insert chart with caption' : 'Insert chart'}
        </Button>
      </div>

      {recent.length > 0 && (
        <div className="chart-recent">
          <h3>Recent charts</h3>
          <ul>
            {recent.map((r) => (
              <li key={r.id}>
                <button type="button" className="chart-recent-item" onClick={() => replace({ ...r.state })} title="Open this chart to change and insert it again">
                  <TypeIcon type={r.state.type} size={18} />
                  <span className="chart-recent-label">{recentLabel(r)}</span>
                  <span className="chart-recent-date">{savedOn(r.savedAt)}</span>
                </button>
              </li>
            ))}
          </ul>
        </div>
      )}
    </div>
  )
}
