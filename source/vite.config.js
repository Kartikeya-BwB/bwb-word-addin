import { defineConfig } from 'vite'
import react from '@vitejs/plugin-react'

// Relative asset paths so the built files work from the GitHub Pages sub-path.
export default defineConfig({
  plugins: [react()],
  base: './',
})
