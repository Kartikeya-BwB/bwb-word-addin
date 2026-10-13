// Options for the chosen chart type: highlight, labels, legend, axis, number format, sort, axis range, gap width,
// series order and colours. Only the options the type uses are shown.

import { ArrowDown, ArrowUp } from 'lucide-react'
import { PALETTE_NAMES } from '../../charts/palette.ts'
import { orderedIndices } from '../../charts/series.ts'
import type { SortOrder, TimeAxis } from '../../charts/spec.ts'
import { Checkbox, SelectField, TextField } from '../components.tsx'
import type { BuilderState, Units } from './builderState.ts'
import { uses } from './builderState.ts'

const UNITS: Array<{ value: Units; label: string }> = [
  { value: '', label: 'None' },
  { value: '%', label: 'Per cent (12%)' },
  { value: 'k', label: 'Thousands (3.5k)' },
  { value: 'm', label: 'Millions (250m)' },
  { value: 'bn', label: 'Billions (1.2bn)' },
  { value: 'tn', label: 'Trillions (1.3tn)' },
]

const DECIMALS = [
  { value: '', label: 'Automatic' },
  { value: '0', label: '0 (12)' },
  { value: '1', label: '1 (12.3)' },
  { value: '2', label: '2 (12.34)' },
  { value: '3', label: '3 (12.345)' },
]

const SORTS: Array<{ value: SortOrder; label: string }> = [
  { value: 'none', label: 'As in the data' },
  { value: 'desc', label: 'Largest first' },
  { value: 'asc', label: 'Smallest first' },
]

const GAPS = [
  { value: '25', label: 'Narrow gaps' },
  { value: '40', label: 'Medium gaps' },
  { value: '60', label: 'Standard gaps' },
  { value: '100', label: 'Wide gaps' },
  { value: '150', label: 'Very wide gaps' },
]

const TIME_AXES: Array<{ value: TimeAxis; label: string }> = [
  { value: 'auto', label: 'Automatic' },
  { value: 'month', label: 'Months' },
  { value: 'quarter', label: 'Quarters' },
  { value: 'year', label: 'Years' },
]

export function ChartOptions({ state, onChange, categories, seriesNames }: {
  state: BuilderState
  onChange: (patch: Partial<BuilderState>) => void
  /** Categories, or Harvey ball options, that can be highlighted. */
  categories: string[]
  seriesNames: string[]
}) {
  const type = state.type
  const single = seriesNames.length <= 1 || type === 'donut' || type === 'harvey'
  const order = orderedIndices(seriesNames.length, state.seriesOrder)
  const move = (k: number, step: number): void => {
    const next = [...order]
    const [moved] = next.splice(k, 1)
    if (moved === undefined) return
    next.splice(Math.min(next.length, Math.max(0, k + step)), 0, moved)
    onChange({ seriesOrder: next })
  }
  const setColour = (index: number, value: string): void => {
    const colours = Array.from({ length: seriesNames.length }, (_, i) => state.colours[i] ?? null)
    colours[index] = Number(value)
    onChange({ colours })
  }
  const highlightOptions = [{ value: '', label: 'None' }, ...categories.filter(Boolean).map((c) => ({ value: c, label: c }))]

  return (
    <div className="chart-options">
      <div className="chart-checks">
        {uses(type, 'labels') && <Checkbox label="Data labels" checked={state.labels} onChange={(labels) => onChange({ labels, ...(labels ? {} : { valueAxis: true }) })} />}
        {uses(type, 'legend') && <Checkbox label="Legend" checked={state.legend} onChange={(legend) => onChange({ legend })} />}
        {uses(type, 'valueAxis') && <Checkbox label="Value axis and gridlines" checked={state.valueAxis} onChange={(valueAxis) => onChange({ valueAxis })} />}
        {uses(type, 'markers') && <Checkbox label="Markers on the line" checked={state.markers} onChange={(markers) => onChange({ markers })} />}
        {uses(type, 'time') && <Checkbox label="Line for today" checked={state.today} onChange={(today) => onChange({ today })} />}
      </div>

      {uses(type, 'highlight') && single && (
        <SelectField label={type === 'harvey' ? 'Highlight an option' : 'Highlight a category'} value={state.highlight}
          options={highlightOptions} onChange={(highlight) => onChange({ highlight })}
          hint={type === 'harvey' ? 'Its balls turn green.' : 'It turns green and the rest navy.'} />
      )}

      {uses(type, 'format') && (
        <>
          <div className="cols-2">
            <SelectField label="Units" value={state.units} options={UNITS} onChange={(units) => onChange({ units })} />
            <TextField label="Currency code" value={state.currency} placeholder="USD"
              onChange={(currency) => onChange({ currency: currency.toUpperCase().replace(/[^A-Z]/g, '').slice(0, 3) })} />
          </div>
          <SelectField label="Decimal places" value={state.decimals} options={DECIMALS} onChange={(decimals) => onChange({ decimals })}
            hint="Labels read like USD 250m: the code, a space, then the number." />
        </>
      )}

      {uses(type, 'sort') && <SelectField label="Order" value={state.sort} options={SORTS} onChange={(sort) => onChange({ sort })} />}

      {uses(type, 'axisRange') && (
        <div className="cols-2">
          <TextField label="Axis minimum" value={state.axisMin} placeholder="Automatic" onChange={(axisMin) => onChange({ axisMin })} />
          <TextField label="Axis maximum" value={state.axisMax} placeholder="Automatic" onChange={(axisMax) => onChange({ axisMax })} />
        </div>
      )}

      {uses(type, 'gap') && (
        <SelectField label="Gap between bars" value={String(state.gapWidth)} options={GAPS}
          onChange={(gap) => onChange({ gapWidth: Number(gap) })} />
      )}

      {uses(type, 'centre') && (
        <TextField label="Centre label" value={state.centreLabel} placeholder="The total" onChange={(centreLabel) => onChange({ centreLabel })}
          hint="Leave empty to show the total." />
      )}

      {uses(type, 'time') && <SelectField label="Time axis" value={state.timeAxis} options={TIME_AXES} onChange={(timeAxis) => onChange({ timeAxis })} />}

      {uses(type, 'series') && seriesNames.length > 1 && (
        <div className="field">
          <span className="field-label">Series order and colours</span>
          <ul className="chart-series">
            {order.map((index, k) => {
              const colour = state.colours[index] ?? index % PALETTE_NAMES.length
              const name = seriesNames[index] ?? `Series ${index + 1}`
              return (
                <li key={index}>
                  <span className={`chart-swatch c${colour}`} aria-hidden />
                  <span className="chart-series-name" title={name}>{name}</span>
                  <select aria-label={`Colour of ${name}`} value={String(colour)} onChange={(e) => setColour(index, e.target.value)}>
                    {PALETTE_NAMES.map((label, i) => <option key={label} value={String(i)}>{label}</option>)}
                  </select>
                  <button type="button" className="chart-icon-btn" aria-label={`Move ${name} up`} disabled={k === 0} onClick={() => move(k, -1)}>
                    <ArrowUp size={14} aria-hidden />
                  </button>
                  <button type="button" className="chart-icon-btn" aria-label={`Move ${name} down`} disabled={k === order.length - 1} onClick={() => move(k, 1)}>
                    <ArrowDown size={14} aria-hidden />
                  </button>
                </li>
              )
            })}
          </ul>
          <span className="field-hint">BwB order is navy, green, green 50, charcoal, light grey.</span>
        </div>
      )}
    </div>
  )
}
