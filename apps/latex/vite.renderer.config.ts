import { defineConfig } from 'vite'

export default defineConfig({
  root: 'src/renderer',
  server: { port: Number(process.env.LATEX_DEV_PORT) || 5177, strictPort: true },
})
