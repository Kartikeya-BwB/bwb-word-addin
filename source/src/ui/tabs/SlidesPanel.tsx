// The PowerPoint Slides tab: a new deck on the BwB template, then the slide tools, with the sticker buttons it shows
// under Stickers.

import { Eraser, FilePlusCorner, Tag } from 'lucide-react'
import { removeAllStickers, setSlideStickers } from '../../ppt/marks.ts'
import { createBwbDeck } from '../../ppt/newDeck.ts'
import { STICKERS } from '../../word/stickers.ts'
import { Button, Section } from '../components.tsx'
import { SlidesTab } from '../ppt/SlidesTab.tsx'
import { useAction } from '../useAction.ts'

function StickerButtons() {
  const { busy, run } = useAction()
  return (
    <>
      <div className="row">
        {STICKERS.map((label) => (
          <Button key={label} small busy={busy === label} disabled={busy !== null} title={`Put “${label}” on the selected slides`}
            onClick={() => void run(label, () => setSlideStickers(label))}>
            <Tag size={12} aria-hidden /> {label}
          </Button>
        ))}
      </div>
      <div className="row" style={{ marginTop: 8 }}>
        <Button small busy={busy === 'remove'} disabled={busy !== null} onClick={() => void run('remove', removeAllStickers)}>
          <Eraser size={12} aria-hidden /> Remove every sticker in the deck
        </Button>
      </div>
      <p className="field-hint">A slide has one sticker at a time: a new one replaces the old. The deck check lists any left before you send.</p>
    </>
  )
}

function NewDeck() {
  const { busy, run } = useAction()
  return (
    <Section id="new" title="New deck" intro="A new presentation on the BwB template: the BwB layouts, colours and fonts, and a title slide to fill in.">
      <Button variant="primary" block busy={busy === 'new'} onClick={() => void run('new', () => createBwbDeck())}>
        <FilePlusCorner size={14} aria-hidden /> Start a new BwB deck
      </Button>
    </Section>
  )
}

export function SlidesPanel() {
  return (
    <>
      <NewDeck />
      <SlidesTab stickers={<StickerButtons />} />
    </>
  )
}
