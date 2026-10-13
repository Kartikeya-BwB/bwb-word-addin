import { afterEach, describe, expect, it } from 'vitest'
import {
  DEFAULT_NOTE, acceptAllChanges, addReviewNote, deleteAllComments, readReviewState, rejectAllChanges, toggleTracking,
} from '../../../src/word/tools/review.ts'
import type { FakeModel } from '../../fakeWord/model.ts'
import { createModel, para } from '../../fakeWord/model.ts'
import { installFakeWord, uninstallFakeWord } from '../../fakeWord/install.ts'

afterEach(uninstallFakeWord)

function reviewed(over: Partial<FakeModel> = {}): FakeModel {
  const p = para('A sentence to review.')
  return createModel([p], { trackedChanges: 3, comments: 2, ...over })
}

describe('review state and tracking', () => {
  it('reads whether changes are tracked and how many changes and comments there are', async () => {
    installFakeWord(reviewed())
    expect(await readReviewState()).toEqual({ tracking: 'Off', changes: 3, comments: 2 })
  })

  it('cannot tell without WordApi 1.4', async () => {
    installFakeWord(reviewed({ apiVersion: 1.3 }))
    expect(await readReviewState()).toEqual({ tracking: null, changes: null, comments: null })
    await expect(toggleTracking()).rejects.toThrow(/Needs a newer Word \(WordApi 1\.4\).*Use Review, Track Changes\./)
  })

  it('turns tracked changes on and off', async () => {
    const model = reviewed()
    installFakeWord(model)
    expect(await toggleTracking()).toBe('Track changes is on: every edit is now recorded.')
    expect(model.changeTrackingMode).toBe('TrackAll')
    expect(await toggleTracking()).toBe('Track changes is off.')
    expect(model.changeTrackingMode).toBe('Off')
  })
})

describe('accepting and rejecting changes', () => {
  it('accepts every tracked change, saying how many', async () => {
    const model = reviewed()
    installFakeWord(model)
    expect(await acceptAllChanges()).toBe('3 tracked changes accepted.')
    expect(model.trackedChanges).toBe(0)
    expect(model.reviewActions).toEqual(['acceptAll'])
    expect(await acceptAllChanges()).toBe('There are no tracked changes to accept.')
  })

  it('rejects every tracked change', async () => {
    const model = reviewed()
    installFakeWord(model)
    expect(await rejectAllChanges()).toBe('3 tracked changes rejected.')
    expect(model.reviewActions).toEqual(['rejectAll'])
  })

  it('uses the desktop call where WordApi 1.6 is missing, and explains when neither is there', async () => {
    const model = reviewed({ apiVersion: 1.5, desktopApiVersion: 1.4 })
    installFakeWord(model)
    expect(await acceptAllChanges()).toBe('Every tracked change accepted.')
    expect(model.reviewActions).toEqual(['acceptAllRevisions'])
    model.desktopApiVersion = 0
    await expect(rejectAllChanges()).rejects.toThrow(/WordApi 1\.6.*Use Review, Reject, Reject All Changes\./)
  })
})

describe('comments', () => {
  it('deletes every comment one by one, or with the desktop call', async () => {
    const model = reviewed({ desktopApiVersion: 0 })
    installFakeWord(model)
    expect(await deleteAllComments()).toBe('2 comments deleted.')
    expect(model.comments).toBe(0)
    expect(await deleteAllComments()).toBe('There are no comments to delete.')
    model.comments = 1
    model.desktopApiVersion = 1.4
    expect(await deleteAllComments()).toBe('1 comment deleted.')
    expect(model.reviewActions).toEqual(['deleteAllComments'])
  })

  it('adds a review note on the selection, or at the cursor', async () => {
    const model = reviewed()
    const p = model.paragraphs[0]
    if (!p) throw new Error('no paragraph')
    installFakeWord(model)
    expect(await addReviewNote('Source needed')).toBe('Review note added at the cursor. Select text first to attach a note to it.')
    model.selectionRange = { paragraph: p, start: 2, end: 10 }
    expect(await addReviewNote('  ')).toBe('Review note added to the selection.')
    expect(model.notes).toEqual(['Source needed', DEFAULT_NOTE])
    expect(model.comments).toBe(4)
  })
})
