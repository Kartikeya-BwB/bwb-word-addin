// The task pane working on a sample held in memory, for trying the pane without Office: `npm run demo`.
// demo.html opens the Word pane on a sample document; demo.html?host=powerpoint opens the PowerPoint pane beside a
// pretend PowerPoint window holding a sample deck.

import { StrictMode, useEffect, useState } from 'react'
import { createRoot } from 'react-dom/client'
import '../../src/taskpane/theme.css'
import './demo.css'
import { App } from '../../src/ui/App.tsx'
import type { FakePresentation, FakeSlide } from '../fakePpt/index.ts'
import { installFakePowerPoint, picture, slide } from '../fakePpt/index.ts'
import { installFakeWord } from '../fakeWord/install.ts'
import { samplePresentation } from './pptSample.ts'
import { sampleModel } from './sample.ts'
import { SlideStage, pictureSources } from './SlideStage.tsx'

const powerpoint = new URLSearchParams(window.location.search).get('host') === 'powerpoint'

/** Calls listeners after anything the add-in does to the deck, so the pretend window redraws. */
const listeners = new Set<() => void>()
const changed = (): void => listeners.forEach((l) => l())

interface LibraryEntry { id: string; title: string; slideId: string; thumb: string }

/** Library slides arrive as their thumbnail on a blank slide: enough to see where they went. */
function libraryInserts(model: FakePresentation): void {
  let index: LibraryEntry[] = []
  fetch('./assets/slides/index.json').then((r) => r.json()).then((d: { slides: LibraryEntry[] }) => { index = d.slides }, () => undefined)
  model.slidesFromBase64 = (_base64, options) => {
    const ids = Array.isArray(options.sourceSlideIds) ? options.sourceSlideIds.map(String) : []
    return ids.map((sid, i): FakeSlide => {
      const entry = index.find((e) => e.slideId === sid)
      const shot = picture({ name: entry?.title ?? 'Library slide', left: 0, top: 0, width: 960, height: 540 })
      if (entry) pictureSources.set(shot.id, `./${entry.thumb}`)
      return slide([shot], { id: `slide-library-${Date.now()}-${i}` })
    })
  }
}

function installPowerPoint(): FakePresentation {
  const model = samplePresentation()
  libraryInserts(model)
  installFakePowerPoint(model)
  const g = globalThis as unknown as { PowerPoint: { run: (b: unknown) => Promise<unknown> }; Office: { context: { document: Record<string, unknown> } } }
  const run = g.PowerPoint.run
  g.PowerPoint.run = async (batch) => {
    try {
      return await run(batch)
    } finally {
      changed()
    }
  }
  const document = g.Office.context.document
  const setSelectedDataAsync = document.setSelectedDataAsync as (...args: unknown[]) => void
  document.setSelectedDataAsync = (...args: unknown[]) => {
    setSelectedDataAsync(...args)
    changed()
  }
  return model
}

function PowerPointDemo({ model }: { model: FakePresentation }) {
  const [, setVersion] = useState(0)
  useEffect(() => {
    const redraw = (): void => setVersion((v) => v + 1)
    listeners.add(redraw)
    return () => void listeners.delete(redraw)
  }, [])
  return (
    <div className="demo-ppt">
      <SlideStage model={model} onChange={changed} />
      <aside className="demo-pane"><App /></aside>
    </div>
  )
}

const root = document.getElementById('root')
if (root) {
  let content
  if (powerpoint) {
    const model = installPowerPoint()
    content = <PowerPointDemo model={model} />
  } else {
    installFakeWord(sampleModel())
    content = <App />
  }
  createRoot(root).render(<StrictMode>{content}</StrictMode>)
}
