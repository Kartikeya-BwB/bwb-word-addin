import type { ButtonHTMLAttributes, CSSProperties, ReactNode } from 'react'
import { useId } from 'react'

export function Section({ id, title, intro, action, weight, children }: {
  id?: string
  title: string
  intro?: ReactNode
  action?: ReactNode
  weight?: 'act'
  children: ReactNode
}) {
  return (
    <section className={`section ${weight ?? ''}`} id={id ? `section-${id}` : undefined}>
      <div className="section-head">
        <h2>{title}</h2>
        {action}
      </div>
      {intro && <p className="section-intro">{intro}</p>}
      {children}
    </section>
  )
}

type ButtonProps = ButtonHTMLAttributes<HTMLButtonElement> & {
  variant?: 'primary' | 'navy' | 'link' | 'plain'
  small?: boolean
  block?: boolean
  busy?: boolean
}

export function Button({ variant = 'plain', small, block, busy, children, className, disabled, ...rest }: ButtonProps) {
  const classes = ['btn', variant !== 'plain' ? variant : '', small ? 'small' : '', block ? 'block' : '', className ?? '']
  return (
    <button type="button" className={classes.filter(Boolean).join(' ')} disabled={disabled || busy} aria-busy={busy || undefined} {...rest}>
      {busy ? 'Working…' : children}
    </button>
  )
}

export function Tile({ mark, label, hint, onClick, disabled, title, markStyle, markClass, preview }: {
  mark: ReactNode
  label: ReactNode
  hint?: ReactNode
  onClick: () => void
  disabled?: boolean
  title?: string
  markStyle?: CSSProperties
  markClass?: string
  preview?: CSSProperties
}) {
  return (
    <button type="button" className="tile" onClick={disabled ? undefined : onClick} aria-disabled={disabled || undefined} title={title}>
      <span className={`tile-mark ${markClass ?? ''}`} style={markStyle} aria-hidden>
        {mark}
      </span>
      <span className="tile-body">
        <span className="tile-label tile-preview" style={preview}>{label}</span>
        {hint && <span className="tile-hint" style={{ display: 'block' }}>{hint}</span>}
      </span>
    </button>
  )
}

export function TextField({ label, value, onChange, hint, placeholder, multiline, type = 'text', required }: {
  label: string
  value: string
  onChange: (value: string) => void
  hint?: string
  placeholder?: string
  multiline?: boolean
  type?: 'text' | 'search' | 'number'
  required?: boolean
}) {
  const id = useId()
  return (
    <div className="field">
      <label htmlFor={id}>{label}{required ? ' (required)' : ''}</label>
      {multiline ? (
        <textarea id={id} value={value} placeholder={placeholder} onChange={(e) => onChange(e.target.value)} />
      ) : (
        <input id={id} type={type} value={value} placeholder={placeholder} required={required} onChange={(e) => onChange(e.target.value)} />
      )}
      {hint && <span className="field-hint">{hint}</span>}
    </div>
  )
}

export function SelectField<T extends string>({ label, value, options, onChange, hint }: {
  label: string
  value: T
  options: Array<{ value: T; label: string }>
  onChange: (value: T) => void
  hint?: string
}) {
  const id = useId()
  return (
    <div className="field">
      <label htmlFor={id}>{label}</label>
      <select id={id} value={value} onChange={(e) => onChange(e.target.value as T)}>
        {options.map((o) => (
          <option key={o.value} value={o.value}>{o.label}</option>
        ))}
      </select>
      {hint && <span className="field-hint">{hint}</span>}
    </div>
  )
}

export function Checkbox({ label, checked, onChange, hint }: {
  label: ReactNode
  checked: boolean
  onChange: (checked: boolean) => void
  hint?: string
}) {
  const id = useId()
  return (
    <div className="check">
      <input id={id} type="checkbox" checked={checked} onChange={(e) => onChange(e.target.checked)} />
      <label htmlFor={id}>
        {label}
        {hint && <span className="field-hint" style={{ display: 'block' }}>{hint}</span>}
      </label>
    </div>
  )
}

export function Notice({ kind = 'ok', children }: { kind?: 'ok' | 'warn' | 'error'; children: ReactNode }) {
  return <div className={`notice ${kind === 'ok' ? '' : kind}`} role={kind === 'error' ? 'alert' : undefined}>{children}</div>
}

/** A collapsible group for a form that most people need only sometimes. */
export function Disclosure({ title, open, children }: { title: string; open?: boolean; children: ReactNode }) {
  return (
    <details className="disclosure" open={open}>
      <summary>{title}</summary>
      <div className="disclosure-body">{children}</div>
    </details>
  )
}
