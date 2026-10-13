import { readdirSync, readFileSync } from 'node:fs'
import { join } from 'node:path'
import { describe, expect, it } from 'vitest'
import { EXPECTED_DIR } from '../../test/helpers/fixtures.ts'
import { LUCIDE_DIR, LUCIDE_NAMES, lucideSvg } from '../../test/helpers/iconRenderer.ts'
import {
  checkColour,
  colourName,
  createIconResolver,
  ICON_COLOURS,
  isIconToken,
  lucideSvgFromNodes,
  type LucideIconNode,
  normaliseIconName,
  parseIconToken,
  pxForCm,
  recolourSvg,
  resolveIconName,
  stripIconTokens,
  suggestIconNames,
} from './iconPolicy.ts'
import { LUCIDE_ALIASES } from './lucideAliases.ts'

interface PolicyReference {
  check_colour: Array<[string, boolean, string | null, string, string | null]>
  resolve: Record<string, string | null>
  suggest3: Record<string, string[]>
  suggest5: Record<string, string[]>
  px_for_cm: Record<string, number>
}

// written by scripts/gen-docx-fixtures.py from the Python icons module
const ref = JSON.parse(readFileSync(join(EXPECTED_DIR, 'icon-policy.json'), 'utf8')) as PolicyReference
const resolve = createIconResolver(LUCIDE_NAMES)
const pool = [...LUCIDE_NAMES, ...Object.keys(LUCIDE_ALIASES)]

describe('colour rules', () => {
  it('substitutes and explains exactly as icons.check_colour', () => {
    for (const [colour, onNavy, fallback, want, note] of ref.check_colour) {
      expect(checkColour(colour, onNavy, fallback), `${colour} ${onNavy} ${fallback}`).toEqual({ colour: want, note })
    }
  })

  it('reads aliases and hex values', () => {
    expect(colourName('Gray')).toBe('charcoal')
    expect(colourName('primary')).toBe('green')
    expect(colourName('#022338')).toBe('navy')
    expect(colourName('constructor')).toBeNull()
    expect(colourName(undefined)).toBeNull()
  })

  it('never puts lime or white on a light background, nor a light colour on navy', () => {
    for (const c of Object.keys(ICON_COLOURS)) {
      expect(['green', 'navy', 'charcoal']).toContain(checkColour(c, false).colour)
      expect(['white', 'lime']).toContain(checkColour(c, true).colour)
    }
  })
})

describe('names', () => {
  it('normalises and resolves older names like icons.resolve', () => {
    for (const [name, want] of Object.entries(ref.resolve)) expect(resolve(name), name).toBe(want)
    expect(normaliseIconName('  Lucide_Leaf Icon ')).toBe('leaf')
    expect(resolveIconName('toString', () => false)).toBeNull()
  })

  it('suggests the same names as icons.suggest', () => {
    for (const [q, want] of Object.entries(ref.suggest3)) expect(suggestIconNames(q, pool, resolve, 3), q).toEqual(want)
    for (const [q, want] of Object.entries(ref.suggest5)) expect(suggestIconNames(q, pool, resolve), q).toEqual(want)
  })

  it('agrees with the alias SVG files lucide-static ships', () => {
    const current = new Set(LUCIDE_NAMES)
    const files = readdirSync(join(LUCIDE_DIR, 'icons')).map((f) => f.replace(/\.svg$/, ''))
    const aliasFiles = files.filter((f) => !current.has(f)).sort()
    expect(aliasFiles).toEqual(Object.keys(LUCIDE_ALIASES).sort())
    for (const [old, now] of Object.entries(LUCIDE_ALIASES)) {
      expect(current.has(now), now).toBe(true)
      expect(lucideSvg(old), old).toBe(lucideSvg(now))
    }
  })
})

describe('svg and sizes', () => {
  it('builds the same SVG from icon nodes as lucide-static ships', () => {
    const nodes = JSON.parse(readFileSync(join(LUCIDE_DIR, 'icon-nodes.json'), 'utf8')) as Record<string, LucideIconNode>
    expect(Object.keys(nodes).length).toBe(LUCIDE_NAMES.length)
    for (const name of LUCIDE_NAMES) expect(lucideSvgFromNodes(nodes[name] ?? []), name).toBe(lucideSvg(name))
  })

  it('recolours the stroke only', () => {
    const svg = recolourSvg(lucideSvg('leaf'), 'green')
    expect(svg).toContain('stroke="#328840"')
    expect(svg).not.toContain('currentColor')
    expect(svg).toContain('fill="none"')
    expect(recolourSvg('<svg stroke="currentColor" stroke-width="2"/>', 'lime', 1.5)).toBe('<svg stroke="#BED62F" stroke-width="1.5"/>')
  })

  it('computes PNG sizes like icons.px_for_cm', () => {
    for (const [cm, px] of Object.entries(ref.px_for_cm)) expect(pxForCm(Number(cm)), cm).toBe(px)
  })
})

describe('{icon:...} tokens', () => {
  it('parses names and colours', () => {
    expect(isIconToken('{icon:leaf}')).toBe(true)
    expect(isIconToken('{icon:leaf:green}')).toBe(true)
    expect(isIconToken('{icon:}')).toBe(false)
    expect(isIconToken('x {icon:leaf}')).toBe(false)
    expect(parseIconToken('{icon:Bar Chart:#328840}')).toEqual({ name: 'Bar Chart', colour: '#328840' })
    expect(parseIconToken('{icon:leaf}')).toEqual({ name: 'leaf', colour: null })
    expect(stripIconTokens('a {icon:leaf} b {icon:x:lime}')).toBe('a  b ')
  })
})
