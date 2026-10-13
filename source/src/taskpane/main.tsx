import { StrictMode } from 'react'
import { createRoot } from 'react-dom/client'
import { App } from '../ui/App.tsx'
import './theme.css'

function render(): void {
  const root = document.getElementById('root')
  if (!root) return
  createRoot(root).render(
    <StrictMode>
      <App />
    </StrictMode>,
  )
}

// Inside Word, wait for Office.js; in a plain browser (npm run dev) render straight away.
if (typeof Office !== 'undefined') void Office.onReady(() => render())
else render()
