// Review tools: tracked changes on or off, accepting or rejecting every change, deleting every comment, and adding a
// review note (a comment) on the selection.

import { UserFacingError, runWord, supports } from '../run.ts'
import { count } from './common.ts'
import { supportsDesktop, unavailable } from './support.ts'

export type TrackingMode = 'Off' | 'TrackAll' | 'TrackMineOnly'

export interface ReviewState {
  /** Null when this Word cannot say (WordApi 1.4 is needed). */
  tracking: TrackingMode | null
  changes: number | null
  comments: number | null
}

export const DEFAULT_NOTE = 'Note for the reviewer: please check this.'

const asMode = (mode: string): TrackingMode => (mode === 'TrackAll' || mode === 'TrackMineOnly' ? mode : 'Off')

async function optional<T>(read: () => Promise<T>): Promise<T | null> {
  try {
    return await read()
  } catch {
    return null
  }
}

/** Whether changes are tracked, and how many tracked changes and comments the document has. */
export async function readReviewState(): Promise<ReviewState> {
  const tracking = supports('1.4')
    ? await optional(() =>
        runWord(async (context) => {
          context.document.load('changeTrackingMode')
          await context.sync()
          return asMode(context.document.changeTrackingMode)
        }),
      )
    : null
  const changes = supports('1.6')
    ? await optional(() =>
        runWord(async (context) => {
          const items = context.document.body.getTrackedChanges()
          items.load('items')
          await context.sync()
          return items.items.length
        }),
      )
    : null
  const comments = supports('1.4')
    ? await optional(() =>
        runWord(async (context) => {
          const items = context.document.body.getComments()
          items.load('items')
          await context.sync()
          return items.items.length
        }),
      )
    : null
  return { tracking, changes, comments }
}

/** Turns tracked changes on, or off when they are on. */
export async function toggleTracking(): Promise<string> {
  const reason = unavailable('tracking')
  if (reason) throw new UserFacingError(`${reason} Use Review, Track Changes.`)
  return runWord(async (context) => {
    context.document.load('changeTrackingMode')
    await context.sync()
    const on = context.document.changeTrackingMode === 'Off'
    context.document.changeTrackingMode = on ? 'TrackAll' : 'Off'
    await context.sync()
    return on ? 'Track changes is on: every edit is now recorded.' : 'Track changes is off.'
  })
}

/** Accepts or rejects every tracked change: in the body with WordApi 1.6, or everywhere with the desktop API. */
async function settleChanges(accept: boolean): Promise<string> {
  const verb = accept ? 'accepted' : 'rejected'
  const reason = unavailable('trackedChanges')
  if (reason) throw new UserFacingError(`${reason} Use Review, ${accept ? 'Accept' : 'Reject'}, ${accept ? 'Accept' : 'Reject'} All Changes.`)
  return runWord(async (context) => {
    if (supports('1.6')) {
      const changes = context.document.body.getTrackedChanges()
      changes.load('items')
      await context.sync()
      const n = changes.items.length
      if (!n) return `There are no tracked changes to ${accept ? 'accept' : 'reject'}.`
      if (accept) changes.acceptAll()
      else changes.rejectAll()
      await context.sync()
      return `${count(n, 'tracked change')} ${verb}.`
    }
    if (!supportsDesktop('1.4')) throw new UserFacingError('This Word cannot do that. Use Review, All Changes.')
    if (accept) context.document.acceptAllRevisions()
    else context.document.rejectAllRevisions()
    await context.sync()
    return `Every tracked change ${verb}.`
  })
}

export const acceptAllChanges = (): Promise<string> => settleChanges(true)
export const rejectAllChanges = (): Promise<string> => settleChanges(false)

/** Deletes every comment, with its replies. */
export async function deleteAllComments(): Promise<string> {
  const reason = unavailable('comments')
  if (reason) throw new UserFacingError(`${reason} Use Review, Delete, Delete All Comments in Document.`)
  return runWord(async (context) => {
    const comments = context.document.body.getComments()
    comments.load('items')
    await context.sync()
    const n = comments.items.length
    if (!n) return 'There are no comments to delete.'
    if (supportsDesktop('1.4')) context.document.deleteAllComments()
    else comments.items.forEach((c) => c.delete())
    await context.sync()
    return `${count(n, 'comment')} deleted.`
  })
}

/** Adds a comment for the reviewer on the selection, or at the cursor. */
export async function addReviewNote(text: string = DEFAULT_NOTE): Promise<string> {
  const reason = unavailable('comments')
  if (reason) throw new UserFacingError(`${reason} Use Review, New Comment.`)
  const note = text.trim() || DEFAULT_NOTE
  return runWord(async (context) => {
    const selection = context.document.getSelection()
    selection.load('text')
    await context.sync()
    selection.insertComment(note)
    await context.sync()
    return selection.text.trim() ? 'Review note added to the selection.' : 'Review note added at the cursor. Select text first to attach a note to it.'
  })
}
