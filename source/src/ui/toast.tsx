import { createContext, useCallback, useContext, useEffect, useRef, useState } from 'react'
import type { ReactNode } from 'react'
import { CircleAlert, CircleCheck } from 'lucide-react'

export type ToastKind = 'ok' | 'error'

interface Toast {
  id: number
  message: string
  kind: ToastKind
}

type Notify = (message: string, kind?: ToastKind) => void

const ToastContext = createContext<Notify>(() => undefined)

export function useNotify(): Notify {
  return useContext(ToastContext)
}

export function ToastProvider({ children }: { children: ReactNode }) {
  const [toast, setToast] = useState<Toast | null>(null)
  const timer = useRef<number | undefined>(undefined)

  const notify = useCallback<Notify>((message, kind = 'ok') => {
    window.clearTimeout(timer.current)
    setToast({ id: Date.now(), message, kind })
    timer.current = window.setTimeout(() => setToast(null), kind === 'error' ? 7000 : 3200)
  }, [])

  useEffect(() => () => window.clearTimeout(timer.current), [])

  return (
    <ToastContext.Provider value={notify}>
      {children}
      <div aria-live="polite" aria-atomic="true">
        {toast && (
          <div key={toast.id} className={`toast ${toast.kind === 'error' ? 'error' : ''}`} role={toast.kind === 'error' ? 'alert' : 'status'}>
            {toast.kind === 'error' ? <CircleAlert size={15} aria-hidden /> : <CircleCheck size={15} aria-hidden />}
            <span>{toast.message}</span>
          </div>
        )}
      </div>
    </ToastContext.Provider>
  )
}
