import { readFile } from 'node:fs/promises'
import { resolve } from 'node:path'
import { describe, expect, it } from 'vitest'
import {
  DEFAULT_OFFICE_BRIDGE_PORTS,
  deploymentConfig,
  deploymentConnectOrigins,
  officeBridgePorts,
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

    expect(manifest).toContain('<Version>0.2.0.0</Version>')
    expect(manifest).toContain(
      '<SourceLocation DefaultValue="https://office.example/taskpane.html?v=0.2.0" />',
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
    { ...validEnv, VITE_WISWORK_PC_BRIDGE_PORTS: '0' },
    { ...validEnv, VITE_WISWORK_PC_BRIDGE_PORTS: '65536' },
    { ...validEnv, VITE_WISWORK_PC_BRIDGE_PORTS: '1.5' },
    { ...validEnv, VITE_WISWORK_PC_BRIDGE_PORTS: '43127,43127' },
    { ...validEnv, VITE_WISWORK_PC_BRIDGE_PORTS: '43127,' },
    { ...validEnv, VITE_WISWORK_PC_BRIDGE_PORTS: '' },
    {
      ...validEnv,
      VITE_WISWORK_PC_BRIDGE_PORTS: Array.from({ length: 129 }, (_, index) =>
        String(10_000 + index),
      ).join(','),
    },
  ])('rejects unsafe deployment configuration', (env) => {
    expect(deploymentConfig(env)).toBeUndefined()
  })

  it('allows only the numeric loopback bridge in taskpane connections', async () => {
    const viteConfig = await readFile(resolve(import.meta.dirname, '../vite.config.ts'), 'utf8')
    const taskpane = await readFile(resolve(import.meta.dirname, '../src/taskpane.html'), 'utf8')
    expect(DEFAULT_OFFICE_BRIDGE_PORTS).toHaveLength(64)
    expect(DEFAULT_OFFICE_BRIDGE_PORTS[0]).toBe(43127)
    expect(new Set(DEFAULT_OFFICE_BRIDGE_PORTS)).toEqual(
      new Set(Array.from({ length: 64 }, (_, index) => 43120 + index)),
    )
    expect(officeBridgePorts({ VITE_WISWORK_PC_BRIDGE_PORTS: '44000,44001' })).toEqual([
      44000, 44001,
    ])
    const origins = deploymentConnectOrigins({}).split(' ')
    expect(origins).toHaveLength(64)
    expect(origins[0]).toBe('http://127.0.0.1:43127')
    expect(origins).toContain('http://127.0.0.1:43120')
    expect(origins.every((origin) => /^http:\/\/127\.0\.0\.1:\d+$/.test(origin))).toBe(true)
    expect(deploymentConnectOrigins({ VITE_WISWORK_PC_BRIDGE_PORTS: '44000,44001' })).toBe(
      'http://127.0.0.1:44000 http://127.0.0.1:44001',
    )
    expect(viteConfig).not.toContain('oauth/callback')
    expect(taskpane).toContain("connect-src 'self' __WISWORK_CONNECT_ORIGINS__")
    expect(taskpane).not.toMatch(/auth\.dev|wisusage|callback/i)
    expect(viteConfig).not.toContain("'Access-Control-Allow-Origin': '*'")
  })

  it('documents the exact PC runtime environment names used by the shell', async () => {
    const readme = await readFile(resolve(import.meta.dirname, '../README.md'), 'utf8')
    expect(readme).toContain('WISWORK_OFFICE_ORIGIN=')
    expect(readme).toContain('WISWORK_OFFICE_BRIDGE_PORTS=')
    expect(readme).not.toContain('WISWORK_OFFICE_ALLOWED_ORIGIN')
  })
})
