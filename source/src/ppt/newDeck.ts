// A new presentation on the BwB template: the BwB master, its seven layouts and the theme, with a title slide and
// one content slide to start from. PowerPoint opens it as a new, unsaved presentation.

import { toBase64 } from '../media/base64.ts'
import { UserFacingError } from '../word/run.ts'
import { pptErrorMessage, requirePpt } from './run.ts'

/** The starter deck that public/templates holds beside the slide library (scripts/make-slide-library.py). */
export const STARTER_DECK = './templates/bwb-deck.pptx'

export async function createBwbDeck(url = STARTER_DECK): Promise<string> {
  requirePpt('1.1', 'Starting a new deck')
  const response = await fetch(url).catch(() => null)
  if (!response?.ok) throw new UserFacingError('The BwB template could not be loaded. Check the connection and try again.')
  const base64 = toBase64(new Uint8Array(await response.arrayBuffer()))
  try {
    await PowerPoint.createPresentation(base64)
  } catch (error) {
    throw new UserFacingError(`PowerPoint could not open the new presentation (${pptErrorMessage(error)}).`, { cause: error })
  }
  return 'A new BwB deck is open in its own window. Save it under its own name.'
}
