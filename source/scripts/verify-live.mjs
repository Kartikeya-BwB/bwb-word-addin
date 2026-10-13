// Checks that everything the published manifest points at is reachable: the pages, the templates, the icon
// library and every ribbon icon.   node scripts/verify-live.mjs [--base https://example.org/path/]
import { readFileSync } from 'node:fs'
import { join } from 'node:path'
import { root } from './ribbon-lib.mjs'

const manifest = readFileSync(join(root, 'public', 'manifest.xml'), 'utf8')
const baseArg = process.argv.indexOf('--base')
const published = /<SourceLocation DefaultValue="([^"]+)index\.html"/.exec(manifest)?.[1]
const base = baseArg > 0 ? process.argv[baseArg + 1].replace(/\/?$/, '/') : published
if (!base) throw new Error('No base URL in the manifest; pass --base.')

const fromManifest = [...manifest.matchAll(/DefaultValue="(https:[^"#]+)/g)].map((m) => m[1]).filter((u) => u.startsWith(published ?? base))
const extra = ['manifest.xml', 'notice.html', 'templates/bwb-base.docx', 'templates/bwb-cover.jpg', 'assets/lucide/index.json', 'assets/brand/logo-white.png']
const urls = [...new Set([...fromManifest.map((u) => base + u.slice((published ?? base).length)), ...extra.map((p) => base + p)])]

let failed = 0
for (let i = 0; i < urls.length; i += 16) {
  const batch = urls.slice(i, i + 16)
  const results = await Promise.all(batch.map(async (url) => {
    try {
      const response = await fetch(url, { method: 'GET' })
      return { url, status: response.status }
    } catch (error) {
      return { url, status: String(error) }
    }
  }))
  for (const r of results) {
    if (r.status !== 200) {
      failed += 1
      console.log(`FAIL ${r.status} ${r.url}`)
    }
  }
}
console.log(failed ? `${failed} of ${urls.length} URLs failed.` : `All ${urls.length} URLs answered 200 at ${base}.`)
process.exitCode = failed ? 1 : 0
