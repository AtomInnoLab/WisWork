import { readFile } from 'node:fs/promises'
import { resolve } from 'node:path'
import { describe, expect, it } from 'vitest'
import {
  deploymentConfig,
  deploymentConnectOrigins,
  renderDeploymentManifest,
} from '../build-config.js'

const manifestPath = resolve(import.meta.dirname, '../public/manifest.xml')
const validEnv = { VITE_WISWORK_ADDIN_ORIGIN: 'https://office.example' }

describe('Office Add-in manifest and routes', () => {
  it('keeps the source manifest explicitly local-development-only', async () => {
    const manifest = await readFile(manifestPath, 'utf8')
    expect(manifest).toContain('DEVELOPMENT-ONLY MANIFEST')
    expect(manifest).toContain('https://localhost:3000/taskpane.html')
  })

  it('renders a configured deployment manifest without localhost or wildcards', async () => {
    const template = await readFile(manifestPath, 'utf8')
    const config = deploymentConfig(validEnv)
    expect(config).toBeDefined()
    const manifest = renderDeploymentManifest(template, config!)

    expect(manifest).toContain(
      '<SourceLocation DefaultValue="https://office.example/taskpane.html" />',
    )
    expect(manifest).toContain('<IconUrl DefaultValue="https://office.example/assets/icon.png" />')
    expect(manifest).toContain('<AppDomain>https://office.example</AppDomain>')
    expect(manifest).not.toContain('auth.example')
    expect(manifest).not.toContain('wisusage.dev.atominnolab.com')
    expect(manifest).not.toContain('localhost')
    expect(manifest).not.toContain('*')
  })

  it.each([
    { VITE_WISWORK_ADDIN_ORIGIN: 'https://office.example/path' },
    { VITE_WISWORK_ADDIN_ORIGIN: 'http://office.example' },
    { VITE_WISWORK_ADDIN_ORIGIN: 'https://*.example' },
  ])('rejects unsafe deployment configuration', (env) => {
    expect(deploymentConfig(env)).toBeUndefined()
  })

  it('allows only the numeric loopback bridge in taskpane connections', async () => {
    const viteConfig = await readFile(resolve(import.meta.dirname, '../vite.config.ts'), 'utf8')
    const taskpane = await readFile(resolve(import.meta.dirname, '../src/taskpane.html'), 'utf8')
    expect(deploymentConnectOrigins()).toBe('http://127.0.0.1:43127')
    expect(viteConfig).not.toContain('oauth/callback')
    expect(taskpane).toContain("connect-src 'self' __WISWORK_CONNECT_ORIGINS__")
    expect(taskpane).not.toMatch(/auth\.dev|wisusage|callback/i)
    expect(viteConfig).not.toContain("'Access-Control-Allow-Origin': '*'")
  })
})
