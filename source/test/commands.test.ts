import { readFileSync } from 'node:fs'
import { join } from 'node:path'
import { afterEach, describe, expect, it } from 'vitest'
import ribbon from '../src/commands/ribbon.json'
import { HANDLERS } from '../src/commands/handlers.ts'
import { createModel, para } from './fakeWord/model.ts'
import { installFakeWord, uninstallFakeWord } from './fakeWord/install.ts'

afterEach(uninstallFakeWord)

interface Entry { id: string; label: string; tip: string; fn?: string; pane?: string; icon: unknown; items?: Entry[] }

const controls = ribbon.groups.flatMap((g) => g.controls as Entry[])
const entries = controls.flatMap((c) => [c, ...(c.items ?? [])])
const TABS = ['start', 'write', 'insert', 'check', 'brand']

describe('the ribbon definition', () => {
  it('has a handler for every function button, and no handler without a button', () => {
    const fns = entries.flatMap((e) => (e.fn ? [e.fn] : []))
    expect(new Set(fns).size).toBe(fns.length)
    expect([...fns].sort()).toEqual(Object.keys(HANDLERS).sort())
  })

  it('opens the pane only at tabs that exist', () => {
    for (const e of entries.filter((x) => x.pane)) expect(TABS, e.id).toContain(e.pane?.split('/')[0])
  })

  it('stays inside Word’s limits and BwB wording rules', () => {
    for (const g of ribbon.groups) expect(g.controls.length, g.id).toBeLessThanOrEqual(6)
    const ids = entries.map((e) => e.id)
    expect(new Set(ids).size).toBe(ids.length)
    for (const e of entries) {
      expect(e.label.length, e.id).toBeLessThanOrEqual(125)
      expect(e.tip.length, e.id).toBeLessThanOrEqual(250)
      expect(`${e.label} ${e.tip}`, e.id).not.toMatch(/[\u2014!]/)
      expect(Boolean(e.fn) !== Boolean(e.pane) || Boolean(e.items), e.id).toBe(true)
    }
  })

  it('uses only Lucide icons that exist', () => {
    const index = JSON.parse(readFileSync(join(import.meta.dirname, '../public/assets/lucide/index.json'), 'utf8')) as { icons: Record<string, unknown> }
    const names = [...ribbon.groups.map((g) => g.icon as unknown), ...entries.map((e) => e.icon)]
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
