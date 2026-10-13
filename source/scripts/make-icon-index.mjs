// Builds public/assets/lucide/index.json from lucide-static: every icon's SVG nodes, its search tags, the
// older names Lucide has renamed, and the BwB topic map. The task pane loads it only when the icon picker opens.
import { mkdirSync, readFileSync, readdirSync, writeFileSync } from 'node:fs'
import { dirname, join } from 'node:path'
import { fileURLToPath } from 'node:url'

const root = join(dirname(fileURLToPath(import.meta.url)), '..')
const lucide = join(root, 'node_modules', 'lucide-static')
const version = JSON.parse(readFileSync(join(lucide, 'package.json'), 'utf8')).version
const nodes = JSON.parse(readFileSync(join(lucide, 'icon-nodes.json'), 'utf8'))
const tags = JSON.parse(readFileSync(join(lucide, 'tags.json'), 'utf8'))
const topics = JSON.parse(readFileSync(join(root, 'src', 'brand', 'iconTopics.json'), 'utf8'))

// An older name ships as its own file with the same drawing as the current icon.
const shape = (file) =>
  readFileSync(join(lucide, 'icons', file), 'utf8')
    .replace(/<!--[\s\S]*?-->/g, '')
    .replace(/class="[^"]*"/, '')
    .replace(/\s+/g, ' ')
    .trim()
const byShape = new Map()
for (const name of Object.keys(nodes)) byShape.set(shape(`${name}.svg`), name)
const aliases = {}
for (const file of readdirSync(join(lucide, 'icons')).sort()) {
  const name = file.replace(/\.svg$/, '')
  if (!file.endsWith('.svg') || nodes[name]) continue
  const target = byShape.get(shape(file))
  if (target) aliases[name] = target
}

const icons = {}
for (const name of Object.keys(nodes).sort()) {
  icons[name] = { n: nodes[name].map(([tag, attrs]) => [tag, attrs]), t: tags[name] ?? [] }
}

const out = join(root, 'public', 'assets', 'lucide')
mkdirSync(out, { recursive: true })
writeFileSync(join(out, 'index.json'), JSON.stringify({ version, licence: 'ISC', icons, aliases, topics }))
writeFileSync(join(out, 'LICENSE.txt'), readFileSync(join(lucide, 'LICENSE'), 'utf8'))
console.log(`lucide ${version}: ${Object.keys(icons).length} icons, ${Object.keys(aliases).length} older names, ${Object.keys(topics).length} topics`)
