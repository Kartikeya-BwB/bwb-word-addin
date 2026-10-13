// Copies the selected slides into a new presentation, the way consultants pull a few slides out for a meeting:
// PowerPoint exports them as a .pptx (API 1.10) and opens it as a new, unsaved presentation.

import { UserFacingError } from '../../word/run.ts'
import { plural } from '../check/util.ts'
import { NEEDS } from './office.ts'
import { pptErrorMessage, requirePpt, runPpt } from '../run.ts'

/** PowerPoint.createPresentation takes at most this many base64 characters. */
const MAX_BASE64 = 71_680_000

export async function copySlidesToNewDeck(): Promise<string> {
  requirePpt(NEEDS.export, 'Copying slides to a new presentation')
  const { base64, count } = await runPpt(async (context) => {
    const selected = context.presentation.getSelectedSlides()
    selected.load('items/id')
    await context.sync()
    const ids = selected.items.map((s) => s.id)
    if (!ids.length) throw new UserFacingError('Select the slides to copy first, in the thumbnails on the left.')
    const exported = context.presentation.slides.exportAsBase64Presentation(ids)
    await context.sync()
    return { base64: exported.value, count: ids.length }
  })
  if (base64.length > MAX_BASE64) {
    throw new UserFacingError('Those slides are too large to copy in one go. Select fewer slides and try again.')
  }
  try {
    await PowerPoint.createPresentation(base64)
  } catch (error) {
    throw new UserFacingError(`PowerPoint could not open the new presentation (${pptErrorMessage(error)}).`, { cause: error })
  }
  return `${plural(count, 'slide')} copied to a new presentation. Save it under its own name.`
}
