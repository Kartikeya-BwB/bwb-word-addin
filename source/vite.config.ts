import { existsSync, readFileSync } from 'node:fs'
import { homedir } from 'node:os'
import { join, resolve } from 'node:path'
import { defineConfig } from 'vite'
import react from '@vitejs/plugin-react'

// Office only loads add-ins over HTTPS. `npx office-addin-dev-certs install` puts a trusted localhost certificate
// here once; the dev server uses it when present (manifest.dev.xml points at https://localhost:3000).
const certs = join(homedir(), '.office-addin-dev-certs')
const https = existsSync(join(certs, 'localhost.crt'))
  ? { cert: readFileSync(join(certs, 'localhost.crt')), key: readFileSync(join(certs, 'localhost.key')) }
  : undefined

// Two pages: the task pane and the hidden page that runs the ribbon commands.
// Relative asset paths so the build works from the GitHub Pages sub-path.
export default defineConfig({
  plugins: [react()],
  base: './',
  build: {
    rolldownOptions: {
      input: {
        taskpane: resolve(import.meta.dirname, 'index.html'),
        commands: resolve(import.meta.dirname, 'commands.html'),
      },
    },
  },
  server: { port: 3000, strictPort: true, ...(https ? { https } : {}) },
})
