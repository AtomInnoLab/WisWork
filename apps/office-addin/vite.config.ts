import { readFile } from 'node:fs/promises'
import { dirname, resolve } from 'node:path'
import { fileURLToPath } from 'node:url'
import react from '@vitejs/plugin-react'
import { getHttpsServerOptions } from 'office-addin-dev-certs'
import { defineConfig, type Plugin } from 'vite'

const here = dirname(fileURLToPath(import.meta.url))

export default defineConfig(async ({ command }) => {
  const icon = await readFile(resolve(here, '../shell/src/main/assets/menu-docx@2x.png'))
  const officeIconPlugin: Plugin = {
    name: 'wiswork-office-icon',
    configureServer(server) {
      server.middlewares.use('/assets/icon.png', (_request, response) => {
        response.setHeader('Content-Type', 'image/png')
        response.end(icon)
      })
    },
    generateBundle() {
      this.emitFile({ type: 'asset', fileName: 'assets/icon.png', source: icon })
    },
  }

  return {
    root: resolve(here, 'src'),
    publicDir: resolve(here, 'public'),
    plugins: [react(), officeIconPlugin],
    build: {
      outDir: resolve(here, 'dist'),
      emptyOutDir: true,
      sourcemap: true,
      rollupOptions: {
        input: resolve(here, 'src/taskpane.html'),
      },
    },
    server: {
      host: 'localhost',
      port: 3000,
      strictPort: true,
      https: command === 'serve' ? await getHttpsServerOptions() : undefined,
      headers: {
        'Access-Control-Allow-Origin': '*',
      },
    },
  }
})
