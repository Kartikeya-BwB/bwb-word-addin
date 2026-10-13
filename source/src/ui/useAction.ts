import { useCallback, useState } from 'react'
import { errorMessage } from '../word/run.ts'
import { useNotify } from './toast.tsx'

/**
 * Runs an async action with a busy flag, shows the success message (if any) and turns any error into a
 * readable toast, so every button in the pane behaves the same way.
 */
export function useAction(): {
  busy: string | null
  run: (key: string, action: () => Promise<string | void>) => Promise<void>
} {
  const notify = useNotify()
  const [busy, setBusy] = useState<string | null>(null)
  const run = useCallback(
    async (key: string, action: () => Promise<string | void>) => {
      setBusy(key)
      try {
        const message = await action()
        if (message) notify(message)
      } catch (error) {
        notify(errorMessage(error), 'error')
      } finally {
        setBusy(null)
      }
    },
    [notify],
  )
  return { busy, run }
}
