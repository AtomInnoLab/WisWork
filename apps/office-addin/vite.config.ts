import { readFile } from 'node:fs/promises'
import { dirname, resolve } from 'node:path'
import { fileURLToPath } from 'node:url'
import react from '@vitejs/plugin-react'
import { getHttpsServerOptions } from 'office-addin-dev-certs'
import { defineConfig, loadEnv, type Plugin } from 'vite'

const here = dirname(fileURLToPath(import.meta.url))
// Keep aligned with WISWORK_MESSAGES_URL; config.ts and transport.ts enforce the same URL at runtime.
const FIXED_MESSAGES_URL = 'https://wisusage.dev.atominnolab.com/v1/messages'

function connectOrigins(env: Record<string, string>): string {
  const origins = new Set<string>()
  for (const key of [
    'VITE_WISWORK_AUTHORIZATION_URL',
    'VITE_WISWORK_TOKEN_URL',
    'VITE_WISWORK_MESSAGES_URL',
  ]) {
    try {
      const url = new URL(env[key] ?? '')
      if (key === 'VITE_WISWORK_MESSAGES_URL' && url.href !== FIXED_MESSAGES_URL) continue
      if (url.protocol === 'https:' && !url.username && !url.password) origins.add(url.origin)
    } catch {
      // Missing or malformed runtime configuration remains unavailable in the app.
    }
  }
  return [...origins].join(' ')
}

export default defineConfig(async ({ command, mode }) => {
  const env = loadEnv(mode, here, '')
  const allowedConnectOrigins = connectOrigins(env)
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
  const securityConfigPlugin: Plugin = {
    name: 'wiswork-office-security-config',
    transformIndexHtml(html) {
      return html.replaceAll('__WISWORK_CONNECT_ORIGINS__', allowedConnectOrigins)
    },
  }

  return {
    root: resolve(here, 'src'),
    envDir: here,
    publicDir: resolve(here, 'public'),
    plugins: [react(), officeIconPlugin, securityConfigPlugin],
    build: {
      outDir: resolve(here, 'dist'),
      emptyOutDir: true,
      sourcemap: true,
      rollupOptions: {
        input: {
          taskpane: resolve(here, 'src/taskpane.html'),
          callback: resolve(here, 'src/oauth/callback.html'),
        },
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
