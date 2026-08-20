import { readFile } from 'node:fs/promises'
import { dirname, resolve } from 'node:path'
import { fileURLToPath } from 'node:url'
import react from '@vitejs/plugin-react'
import { getHttpsServerOptions } from 'office-addin-dev-certs'
import { defineConfig, loadEnv, type Plugin } from 'vite'
import {
  deploymentConfig,
  deploymentConnectOrigins,
  renderDeploymentManifest,
  rewriteOAuthCallbackRequest,
} from './build-config.js'

const here = dirname(fileURLToPath(import.meta.url))
export default defineConfig(async ({ command, mode }) => {
  const env = loadEnv(mode, here, '')
  const deployment = deploymentConfig(env)
  const allowedConnectOrigins = deployment ? deploymentConnectOrigins(deployment) : ''
  const icon = await readFile(resolve(here, '../shell/src/main/assets/menu-docx@2x.png'))
  const manifestTemplate = await readFile(resolve(here, 'public/manifest.xml'), 'utf8')
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
    enforce: 'post',
    transformIndexHtml(html) {
      return html.replaceAll('__WISWORK_CONNECT_ORIGINS__', allowedConnectOrigins)
    },
    configureServer(server) {
      server.middlewares.use((request, _response, next) => {
        request.url = rewriteOAuthCallbackRequest(request.url)
        next()
      })
    },
    configurePreviewServer(server) {
      server.middlewares.use((request, _response, next) => {
        request.url = rewriteOAuthCallbackRequest(request.url)
        next()
      })
    },
    generateBundle(_options, bundle) {
      const callback = Object.values(bundle).find(
        (entry) => entry.type === 'asset' && entry.fileName === 'oauth/callback.html',
      )
      if (!callback || callback.type !== 'asset') this.error('OAuth callback HTML was not emitted')
      this.emitFile({ type: 'asset', fileName: 'oauth/callback', source: callback.source })
      if (deployment) {
        this.emitFile({
          type: 'asset',
          fileName: 'manifest.xml',
          source: renderDeploymentManifest(manifestTemplate, deployment),
        })
      }
    },
  }

  return {
    root: resolve(here, 'src'),
    envDir: here,
    publicDir: false as const,
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
    },
  }
})
