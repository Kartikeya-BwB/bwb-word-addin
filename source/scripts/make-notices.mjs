// Writes public/THIRD-PARTY-NOTICES.txt: every third-party library bundled into the add-in, and every asset library it
// ships, with its licence text, as their licences require. Run by `npm run assets`; the pane's Brand tab links to it.

import { readFileSync, writeFileSync } from 'node:fs'
import { join } from 'node:path'
import { root } from './ribbon-lib.mjs'

const modules = join(root, 'node_modules')
const version = (name) => JSON.parse(readFileSync(join(modules, name, 'package.json'), 'utf8')).version

/** What the add-in ships, and where each licence text comes from. */
const ENTRIES = [
  { name: 'React and React DOM', what: 'the task pane (bundled)', pkg: 'react', licence: 'MIT', file: join(modules, 'react', 'LICENSE') },
  { name: 'fflate', what: 'reading and writing .docx and .pptx files (bundled)', pkg: 'fflate', licence: 'MIT', file: join(modules, 'fflate', 'LICENSE') },
  { name: 'Lucide', what: 'icons in the pane and the icon library', pkg: 'lucide-react', licence: 'ISC', file: join(root, 'public/assets/lucide/LICENSE.txt') },
  { name: 'Tabler Icons', what: 'the Tabler icon library', pkg: '@tabler/icons', licence: 'MIT', file: join(root, 'public/assets/icons/LICENSE-tabler.txt') },
  { name: 'flag-icons', what: 'flags', pkg: 'flag-icons', licence: 'MIT', file: join(root, 'public/assets/icons/flags/LICENSE.txt') },
  {
    name: 'world-atlas, with Natural Earth data', what: 'country and region maps (Natural Earth data is in the public domain)', pkg: 'world-atlas',
    licence: 'ISC', file: join(root, 'public/assets/icons/maps/LICENSE.txt'),
  },
]

const rule = '-'.repeat(78)
const parts = [
  'BwB Toolkit: third-party notices',
  '',
  'The BwB Toolkit for Word and PowerPoint includes the following third-party software and assets. Each is used under',
  'its licence, reproduced below. Fonts are not included: Candara and Aptos come with Microsoft 365.',
  '',
  ...ENTRIES.map((e) => `- ${e.name} ${version(e.pkg)} (${e.licence}): ${e.what}`),
  '',
]
for (const e of ENTRIES) {
  parts.push(rule, `${e.name} ${version(e.pkg)}, ${e.licence}`, rule, '', readFileSync(e.file, 'utf8').trim(), '')
}
writeFileSync(join(root, 'public', 'THIRD-PARTY-NOTICES.txt'), `${parts.join('\n')}\n`)
console.log(`Third-party notices written for ${ENTRIES.length} libraries.`)
