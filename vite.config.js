import { fileURLToPath } from 'url'
import { defineConfig } from 'vite'
import react from '@vitejs/plugin-react'

// https://vite.dev/config/
export default defineConfig({
  plugins: [react()],
  resolve: {
    alias: {
      'remark-math': fileURLToPath(
        new URL('./node_modules/remark-math/index.js', import.meta.url),
      ),
      'rehype-katex': fileURLToPath(
        new URL('./node_modules/rehype-katex/index.js', import.meta.url),
      ),
    },
  },
  optimizeDeps: {
    include: ['remark-math', 'rehype-katex', 'katex'],
  },
  server: {
    host: true,
  },
})
