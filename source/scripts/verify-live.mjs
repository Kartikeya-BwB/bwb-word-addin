// Compares the fresh build in dist/ with the files GitHub Pages serves from the repo root.
// Every built file must be byte-identical to its deployed copy; exits 1 otherwise.
import { createHash } from 'node:crypto'
import { existsSync, readFileSync, readdirSync } from 'node:fs'
import { dirname, join, relative, sep } from 'node:path'
import { fileURLToPath } from 'node:url'

const sourceDir = join(dirname(fileURLToPath(import.meta.url)), '..')
const distDir = join(sourceDir, 'dist')
const liveDir = join(sourceDir, '..')

const sha256 = (file) => createHash('sha256').update(readFileSync(file)).digest('hex')

const listFiles = (dir) =>
  readdirSync(dir, { withFileTypes: true }).flatMap((entry) => {
    const path = join(dir, entry.name)
    return entry.isDirectory() ? listFiles(path) : [path]
  })

let mismatches = 0
for (const built of listFiles(distDir).sort()) {
  const rel = relative(distDir, built).split(sep).join('/')
  const live = join(liveDir, rel)
  if (!existsSync(live)) {
    console.log(`NOT DEPLOYED  ${rel}`)
    mismatches++
  } else if (sha256(built) !== sha256(live)) {
    console.log(`DIFFERENT     ${rel}`)
    mismatches++
  } else {
    console.log(`identical     ${rel}  ${sha256(built).slice(0, 16)}`)
  }
}

if (mismatches) {
  console.log(`\n${mismatches} file(s) differ from what is deployed at the repo root.`)
  process.exit(1)
}
console.log('\nThe build is byte-identical to the deployed add-in.')
