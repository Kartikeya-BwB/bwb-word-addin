import { useState } from 'react'
import { altTextFromCaptions, captionAllPictures, captionSelectedPicture } from '../../word/tools/captions.ts'
import { TEXT_WIDTH_CM } from '../../word/tools/common.ts'
import type { PictureAlignment } from '../../word/tools/pictures.ts'
import { alignPictures, matchPictureWidths, setPictureWidth } from '../../word/tools/pictures.ts'
import { unavailable } from '../../word/tools/support.ts'
import { Disclosure, Section, TextField } from '../components.tsx'
import { useAction } from '../useAction.ts'
import { ToolButton } from './ToolButton.tsx'

const ALIGNMENTS: Array<{ value: PictureAlignment; label: string }> = [
  { value: 'Left', label: 'Left' },
  { value: 'Centered', label: 'Centre' },
  { value: 'Right', label: 'Right' },
]

export function PictureTools() {
  const { busy, run } = useAction()
  const working = busy !== null
  const reason = unavailable('core')
  const [width, setWidth] = useState('12')
  const [title, setTitle] = useState('')
  const [source, setSource] = useState('')

  const tool = (key: string, label: string, action: () => Promise<string>, variant?: 'primary' | 'navy') => (
    <ToolButton key={key} label={label} busy={busy === key} working={working} reason={reason} onClick={() => void run(key, action)}
      {...(variant ? { variant } : {})} />
  )

  return (
    <Section id="pictures" title="Pictures"
      intro={'Select pictures, or put the cursor on a picture\u2019s line. With nothing selected, these change every figure and leave icons and logos alone.'}>
      <div className="tool-row">
        {tool('fit', `Fit to text width (${TEXT_WIDTH_CM} cm)`, () => setPictureWidth('text'), 'primary')}
        {tool('same', 'Same width as the first', matchPictureWidths)}
      </div>
      <div className="tool-label">Align on the line</div>
      <div className="tool-row">
        {ALIGNMENTS.map((a) => tool(`align-${a.value}`, a.label, () => alignPictures(a.value)))}
      </div>
      <Disclosure title="Exact width">
        <TextField label="Width in cm" type="number" value={width} onChange={setWidth}
          hint={`Proportions are kept. The BwB text width is ${TEXT_WIDTH_CM} cm.`} />
        {tool('width', 'Set the width', () => setPictureWidth(Number(width)), 'navy')}
      </Disclosure>
      <Disclosure title="Captions and alt text">
        <TextField label="Caption title" value={title} onChange={setTitle} placeholder="Climate finance flows by region, 2025"
          hint="Leave empty for grey placeholders to type over. Numbered automatically: Figure 1, Figure 2." />
        <TextField label="Source" value={source} onChange={setSource} placeholder="OECD (2026)" />
        <div className="tool-row">
          {tool('caption', 'Add a caption below the picture', async () => {
            const message = await captionSelectedPicture(title, source)
            setTitle('')
            setSource('')
            return message
          }, 'primary')}
        </div>
        <div className="tool-row">
          {tool('caption-all', 'Caption every figure that has none', captionAllPictures)}
          {tool('alt', 'Alt text from captions', altTextFromCaptions)}
        </div>
      </Disclosure>
    </Section>
  )
}
