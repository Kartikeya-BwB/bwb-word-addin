// The task pane working on a sample document held in memory, for trying the pane without Word: `npm run demo`.
import { StrictMode } from 'react'
import { createRoot } from 'react-dom/client'
import '../../src/taskpane/theme.css'
import { App } from '../../src/ui/App.tsx'
import { installFakeWord } from '../fakeWord/install.ts'
import { sampleModel } from './sample.ts'

installFakeWord(sampleModel())
const root = document.getElementById('root')
if (root) {
  createRoot(root).render(
    <StrictMode>
      <App />
    </StrictMode>,
  )
}
