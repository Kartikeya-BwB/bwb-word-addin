import { readFileSync } from 'node:fs'
import { join } from 'node:path'
import { afterEach, describe, expect, it } from 'vitest'
import wordRibbon from '../src/commands/ribbon-word.json'
import pptRibbon from '../src/commands/ribbon-powerpoint.json'
import { HANDLERS } from '../src/commands/handlers.ts'
import { createModel, para } from './fakeWord/model.ts'
import { installFakeWord, uninstallFakeWord } from './fakeWord/install.ts'

afterEach(uninstallFakeWord)

interface Entry { id: string; label: string; tip: string; fn?: string; pane?: string; icon: unknown; items?: Entry[] }
interface Ribbon { groups: Array<{ id: string; icon: unknown; controls: Entry[] }> }

const RIBBONS: Array<{ host: string; ribbon: Ribbon; tabs: string[] }> = [
  { host: 'Word', ribbon: wordRibbon as Ribbon, tabs: ['start', 'write', 'insert', 'tools', 'check', 'brand'] },
  { host: 'PowerPoint', ribbon: pptRibbon as Ribbon, tabs: ['slides', 'arrange', 'insert', 'check', 'brand'] },
]
const entriesOf = (r: Ribbon): Entry[] => r.groups.flatMap((g) => g.controls).flatMap((c) => [c, ...(c.items ?? [])])

describe('the ribbon definitions', () => {
  it('have a handler for every function button, and no handler without a button', () => {
    const fns = RIBBONS.flatMap(({ ribbon }) => entriesOf(ribbon).flatMap((e) => (e.fn ? [e.fn] : [])))
    expect(new Set(fns).size).toBe(fns.length)
    expect([...fns].sort()).toEqual(Object.keys(HANDLERS).sort())
  })

  it.each(RIBBONS)('$host: opens the pane only at tabs that exist', ({ ribbon, tabs }) => {
    for (const e of entriesOf(ribbon).filter((x) => x.pane)) expect(tabs, e.id).toContain(e.pane?.split('/')[0])
  })

  it.each(RIBBONS)('$host: stays inside Office\u2019s limits and BwB wording rules', ({ ribbon }) => {
    for (const g of ribbon.groups) expect(g.controls.length, g.id).toBeLessThanOrEqual(6)
    const entries = entriesOf(ribbon)
    const ids = entries.map((e) => e.id)
    expect(new Set(ids).size).toBe(ids.length)
    for (const e of entries) {
      expect(e.label.length, e.id).toBeLessThanOrEqual(125)
      expect(e.tip.length, e.id).toBeLessThanOrEqual(250)
      expect(`${e.label} ${e.tip}`, e.id).not.toMatch(/[\u2014!]/)
      expect(Boolean(e.fn) !== Boolean(e.pane) || Boolean(e.items), e.id).toBe(true)
      if (e.items) expect(e.items.length, e.id).toBeLessThanOrEqual(15)
    }
  })

  it.each(RIBBONS)('$host: uses only Lucide icons that exist', ({ ribbon }) => {
    const index = JSON.parse(readFileSync(join(import.meta.dirname, '../public/assets/lucide/index.json'), 'utf8')) as { icons: Record<string, unknown> }
    const names = [...ribbon.groups.map((g) => g.icon), ...entriesOf(ribbon).map((e) => e.icon)]
      .map((i) => (typeof i === 'string' ? i : i && typeof i === 'object' && 'name' in i ? String(i.name) : null))
      .filter((n): n is string => n !== null)
    for (const n of names) expect(index.icons[n], n).toBeDefined()
  })
})

describe('ribbon handlers in Word', () => {
  it('apply styles to the selection', async () => {
    const p = para('Background')
    installFakeWord(createModel([p]))
    await HANDLERS.styleHeading2?.()
    expect(p.styleBuiltIn).toBe('Heading2')
  })

  it('set and remove stickers', async () => {
    const model = createModel([para('Body')])
    installFakeWord(model)
    await HANDLERS.stickerForDiscussion?.()
    expect(model.sections[0]?.headers.Primary.controls[0]?.paragraphs[0]?.text).toBe('FOR DISCUSSION')
    await HANDLERS.stickerRemove?.()
    expect(model.sections[0]?.headers.Primary.controls).toHaveLength(0)
  })

  it('insert library blocks below the cursor', async () => {
    const model = createModel([para('An opening paragraph.')])
    installFakeWord(model)
    await HANDLERS.blockCallout?.()
    expect(model.inserts).toHaveLength(1)
    expect(model.inserts[0]?.location).toBe('After')
  })
})
