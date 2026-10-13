// Copies the build (dist/) to the repository root, which GitHub Pages serves. Whatever the previous deploy put
// there (listed in .deployed.json) is removed first, so nothing stale is left; nothing else at the root is touched.
import { cpSync, existsSync, readFileSync, readdirSync, rmSync, writeFileSync } from 'node:fs'
import { join } from 'node:path'
import { root } from './ribbon-lib.mjs'

const repo = join(root, '..')
const dist = join(root, 'dist')
const record = join(repo, '.deployed.json')
// The first version's build output, from before deploys were recorded.
const FIRST_VERSION = ['index.html', 'assets', 'manifest.xml', 'favicon.svg', 'icons.svg', 'icon-16.png', 'icon-32.png', 'icon-80.png']
const NEVER = new Set(['.git', '.github', 'source', 'README.md', 'HOW-TO-DEPLOY.md', '.deployed.json', '.nojekyll', 'REA-FINDINGS.md'])

if (!existsSync(join(dist, 'index.html'))) throw new Error('Run the build first: dist/index.html is missing.')

const previous = existsSync(record) ? JSON.parse(readFileSync(record, 'utf8')).files : FIRST_VERSION
for (const name of previous) if (!NEVER.has(name)) rmSync(join(repo, name), { recursive: true, force: true })

const files = readdirSync(dist).filter((name) => !NEVER.has(name)).sort()
for (const name of files) cpSync(join(dist, name), join(repo, name), { recursive: true })
writeFileSync(join(repo, '.nojekyll'), '')
writeFileSync(record, `${JSON.stringify({ note: 'Written by source/scripts/deploy-to-root.mjs', files }, null, 2)}\n`)
console.log(`Deployed ${files.length} entries to the repository root: ${files.join(', ')}`)
