import type { ReactNode } from 'react'
import { Button } from '../components.tsx'

/**
 * A pane button for a Word tool. When this Word cannot run the tool, the button is disabled and the reason shows
 * under it, so nobody is left wondering why it does nothing.
 */
export function ToolButton({ label, onClick, busy, working, reason, variant, block, small, title, children }: {
  label: string
  onClick: () => void
  busy?: boolean
  /** Another tool is running. */
  working?: boolean
  /** Why this Word cannot run the tool, or null when it can. */
  reason?: string | null
  variant?: 'primary' | 'navy' | 'plain'
  block?: boolean
  small?: boolean
  title?: string
  children?: ReactNode
}) {
  const button = (
    <Button
      variant={variant ?? 'plain'}
      block={block ?? false}
      small={small ?? false}
      busy={busy ?? false}
      disabled={Boolean(reason) || (working ?? false)}
      title={reason ?? title}
      onClick={onClick}
    >
      {children ?? label}
    </Button>
  )
  if (!reason) return button
  return (
    <span className="tool-unavailable">
      {button}
      <span className="field-hint">{reason}</span>
    </span>
  )
}
