import { useState } from 'react'
import { hasPowerPoint, supportsPpt, updateMessage } from '../../../ppt/run.ts'
import type { Focus } from '../../../ppt/tools/frame.ts'
import { fillWithPicture } from '../../../ppt/tools/frame.ts'
import { ImagePicker } from '../../library/ImagePicker.tsx'
import { Button, Section, SelectField } from '../../components.tsx'
import { useAction } from '../../useAction.ts'
import './tools.css'

const FOCUS: Array<{ value: Focus; label: string }> = [
  { value: 'top', label: 'The top, for faces' },
  { value: 'centre', label: 'The middle' },
  { value: 'bottom', label: 'The bottom' },
]

/** Fills the selected shape with a BwB library photo, cropped to the shape rather than stretched. */
export function PictureFrameSection() {
  const { run } = useAction()
  const [focus, setFocus] = useState<Focus>('centre')
  const [open, setOpen] = useState(false)
  const blocked = hasPowerPoint() && !supportsPpt('1.8')
  return (
    <Section id="picture-frame" title="Picture into frame"
      intro="Select a shape, such as a circle for a portrait, then choose a photo. It is cropped to the shape, never stretched.">
      {blocked && <p className="tools-note">{updateMessage('Filling a shape with a photo')}</p>}
      <SelectField label="When the photo is taller than the shape, keep" value={focus} options={FOCUS} onChange={setFocus} />
      {open
        ? <ImagePicker chooseLabel="Fill the shape with" onChoose={(image) => void run('frame', () => fillWithPicture(image, focus))} />
        : <Button disabled={blocked} onClick={() => setOpen(true)}>Choose a photo</Button>}
    </Section>
  )
}
