import { afterEach, describe, expect, it } from 'vitest'
import wordRibbon from '../../../src/commands/ribbon-word.json'
import { WORD_TOOL_HANDLERS } from '../../../src/word/tools/handlers.ts'
import { createModel, para } from '../../fakeWord/model.ts'
import { installFakeWord, uninstallFakeWord } from '../../fakeWord/install.ts'

afterEach(uninstallFakeWord)

interface Entry { fn?: string; items?: Entry[] }
interface Ribbon { groups: Array<{ id: string; controls: Entry[] }> }

describe('WORD_TOOL_HANDLERS', () => {
  it('has exactly the functions of the ribbon\u2019s Tools group', () => {
    const tools = (wordRibbon as Ribbon).groups.find((g) => g.id === 'Tools')
    const fns = (tools?.controls ?? []).flatMap((c) => [c, ...(c.items ?? [])]).flatMap((e) => (e.fn ? [e.fn] : []))
    expect(fns).toHaveLength(23)
    expect(Object.keys(WORD_TOOL_HANDLERS).sort()).toEqual([...fns].sort())
  })

  it('run against Word', async () => {
    const p = para('blended finance')
    const model = createModel([p], { selectionRange: { paragraph: p, start: 0, end: 15 }, comments: 1 })
    installFakeWord(model)
    await WORD_TOOL_HANDLERS.wordCaseTitle?.()
    await WORD_TOOL_HANDLERS.wordToggleTracking?.()
    await WORD_TOOL_HANDLERS.wordAddReviewNote?.()
    await WORD_TOOL_HANDLERS.wordDeleteAllComments?.()
    expect(p.text).toBe('Blended Finance')
    expect(model.changeTrackingMode).toBe('TrackAll')
    expect(model.notes).toEqual(['Note for the reviewer: please check this.'])
    expect(model.comments).toBe(0)
  })
})
