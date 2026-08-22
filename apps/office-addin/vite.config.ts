import { readFile } from 'node:fs/promises'
import { execFileSync } from 'node:child_process'
import { dirname, resolve } from 'node:path'
import { fileURLToPath } from 'node:url'
import react from '@vitejs/plugin-react'
import { getHttpsServerOptions } from 'office-addin-dev-certs'
import { defineConfig, loadEnv, type Plugin } from 'vite'
import {
  deploymentConfig,
  deploymentConnectOrigins,
  officeBuildId,
  renderDeploymentManifest,
} from './build-config.js'

const here = dirname(fileURLToPath(import.meta.url))
function gitBuildId(): string {
  try {
    return execFileSync('git', ['rev-parse', '--short=12', 'HEAD'], {
      cwd: here,
      encoding: 'utf8',
      stdio: ['ignore', 'pipe', 'ignore'],
    }).trim()
  } catch {
    return 'development'
  }
}
export default defineConfig(async ({ command, mode }) => {
  const env = loadEnv(mode, here, '')
  const buildId = officeBuildId(env, process.env.GITHUB_SHA?.slice(0, 12) || gitBuildId())
  const deployment = deploymentConfig(env)
  const allowedConnectOrigins = deploymentConnectOrigins(env)
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
    generateBundle() {
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
    define: {
      __WISWORK_OFFICE_BUILD_ID__: JSON.stringify(buildId),
    },
    worker: {
      format: 'es' as const,
    },
    build: {
      outDir: resolve(here, 'dist'),
      emptyOutDir: true,
      sourcemap: false,
      rollupOptions: {
        input: {
          taskpane: resolve(here, 'src/taskpane.html'),
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
