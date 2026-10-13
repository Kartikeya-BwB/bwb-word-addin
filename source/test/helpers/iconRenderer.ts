/**
 * Node-only IconRenderer for tests: Lucide SVGs from lucide-static (the same 1.52.0 set as the BwB skill) recoloured
 * with iconPolicy, drawn to PNG with @resvg/resvg-js.
 */
import { readFileSync } from 'node:fs'
import { join } from 'node:path'
import { fileURLToPath } from 'node:url'
import { Resvg } from '@resvg/resvg-js'
import { createIconResolver, recolourSvg } from '../../src/docx/iconPolicy.ts'
import { LUCIDE_ALIASES } from '../../src/docx/lucideAliases.ts'
import type { IconColour, IconRenderer } from '../../src/docx/types.ts'

export const LUCIDE_DIR = fileURLToPath(new URL('../../node_modules/lucide-static/', import.meta.url))

const tags = JSON.parse(readFileSync(join(LUCIDE_DIR, 'tags.json'), 'utf8')) as Record<string, string[]>
/** Current Lucide names, in the order of lucide-static's tags.json (the skill's index uses the same order). */
export const LUCIDE_NAMES: readonly string[] = Object.keys(tags)
const resolveName = createIconResolver(LUCIDE_NAMES)

/** A lucide-static SVG without its licence comment, class attribute and indentation. */
export function cleanLucideSvg(text: string): string {
  return text
    .replace(/<!--[\s\S]*?-->/g, '')
    .replace(/\s*class="[^"]*"/, '')
    .replace(/\s+/g, ' ')
    .replace(/>\s+</g, '><')
    .replace(/\s*\/>/g, '/>')
    .replace(/<svg (.*?)\s*>/, '<svg $1>')
    .trim()
}

export function lucideSvg(name: string): string {
  return cleanLucideSvg(readFileSync(join(LUCIDE_DIR, 'icons', `${name}.svg`), 'utf8'))
}

export interface TestIconRenderer extends IconRenderer {
  /** Number of PNG renderings made, to check that each icon and colour is drawn once. */
  pngCalls: number
}

export function createTestIconRenderer(options: { failPng?: boolean } = {}): TestIconRenderer {
  const cache = new Map<string, Uint8Array>()
  const r: TestIconRenderer = {
    pngCalls: 0,
    resolve: (name) => resolveName(name),
    svg: (name: string, colour: IconColour) => {
      const n = resolveName(name)
      return n ? recolourSvg(lucideSvg(n), colour) : null
    },
    png: async (name: string, colour: IconColour, px: number) => {
      r.pngCalls += 1
      if (options.failPng) return null
      const key = `${name}|${colour}|${px}`
      const hit = cache.get(key)
      if (hit) return hit
      const svg = r.svg(name, colour)
      if (!svg) return null
      const png = new Uint8Array(new Resvg(svg, { fitTo: { mode: 'width', value: px }, font: { loadSystemFonts: false } }).render().asPng())
      cache.set(key, png)
      return png
    },
    names: () => [...LUCIDE_NAMES, ...Object.keys(LUCIDE_ALIASES)],
  }
  return r
}
