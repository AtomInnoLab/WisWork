import { readFile } from 'node:fs/promises'
import { resolve } from 'node:path'
import { describe, expect, it } from 'vitest'
import {
  DEFAULT_OFFICE_BRIDGE_PORTS,
  deploymentConfig,
  deploymentConnectOrigins,
  officeBridgePorts,
  officeBuildId,
  officeCapabilityFlags,
  officeRemoteDiagnosticsEnabled,
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

    expect(manifest).toContain('<Version>0.3.3.0</Version>')
    expect(manifest).toContain(
      '<SourceLocation DefaultValue="https://office.example/taskpane.html?v=0.3.3" />',
    )
    expect(manifest).toContain('<IconUrl DefaultValue="https://office.example/assets/icon.png" />')
    expect(manifest).toContain('<AppDomain>https://office.example</AppDomain>')
    expect(manifest).not.toContain('auth.example')
    expect(manifest).not.toContain('wisusage.atominnolab.com')
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

  it('uses only the fixed WSS relay by default and loopback only in rollback mode', async () => {
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
    expect(deploymentConnectOrigins({})).toBe('wss://office.8-216-134-194.sslip.io')
    expect(
      deploymentConnectOrigins({
        VITE_WISWORK_OFFICE_TRANSPORT: 'loopback',
        VITE_WISWORK_PC_BRIDGE_PORTS: '44000,44001',
      }),
    ).toBe('http://127.0.0.1:44000 http://127.0.0.1:44001')
    expect(viteConfig).not.toContain('oauth/callback')
    expect(taskpane).toContain("connect-src 'self' __WISWORK_CONNECT_ORIGINS__")
    expect(taskpane).not.toMatch(/auth\.dev|wisusage|callback/i)
    expect(viteConfig).not.toContain("'Access-Control-Allow-Origin': '*'")
  })

  it('parses independent exact capability rollback flags and rejects invalid values', () => {
    expect(officeCapabilityFlags({})).toEqual({
      conversions: true,
      skillPackages: true,
      importMedia: true,
    })
    expect(
      officeCapabilityFlags({
        VITE_WISWORK_OFFICE_CONVERSIONS: '0',
        VITE_WISWORK_OFFICE_SKILL_PACKAGES: '0',
        VITE_WISWORK_OFFICE_IMPORT_MEDIA: '0',
      }),
    ).toEqual({ conversions: false, skillPackages: false, importMedia: false })
    expect(() => officeCapabilityFlags({ VITE_WISWORK_OFFICE_CONVERSIONS: 'false' })).toThrow(
      'invalid_office_capability_flags',
    )
    expect(
      deploymentConfig({ ...validEnv, VITE_WISWORK_OFFICE_SKILL_PACKAGES: 'false' }),
    ).toBeUndefined()
  })

  it('enables safe remote diagnostics by default with an exact rollback flag', () => {
    expect(officeRemoteDiagnosticsEnabled({})).toBe(true)
    expect(officeRemoteDiagnosticsEnabled({ VITE_WISWORK_OFFICE_REMOTE_DIAGNOSTICS: '1' })).toBe(
      true,
    )
    expect(officeRemoteDiagnosticsEnabled({ VITE_WISWORK_OFFICE_REMOTE_DIAGNOSTICS: '0' })).toBe(
      false,
    )
    expect(() =>
      officeRemoteDiagnosticsEnabled({ VITE_WISWORK_OFFICE_REMOTE_DIAGNOSTICS: 'true' }),
    ).toThrow('invalid_office_remote_diagnostics')
    expect(
      deploymentConfig({ ...validEnv, VITE_WISWORK_OFFICE_REMOTE_DIAGNOSTICS: 'true' }),
    ).toBeUndefined()
  })

  it('uses a validated deploy build identifier instead of an uncorrelated unknown value', async () => {
    expect(officeBuildId({ VITE_WISWORK_OFFICE_BUILD_ID: 'abc123def456' }, 'fallback')).toBe(
      'abc123def456',
    )
    expect(officeBuildId({}, 'abc123def456')).toBe('abc123def456')
    expect(() =>
      officeBuildId({ VITE_WISWORK_OFFICE_BUILD_ID: 'bad build id' }, 'fallback'),
    ).toThrow('invalid_office_build_id')
    const viteConfig = await readFile(resolve(import.meta.dirname, '../vite.config.ts'), 'utf8')
    expect(viteConfig).toContain('__WISWORK_OFFICE_BUILD_ID__')
    expect(viteConfig).toContain("rev-parse', '--short=12', 'HEAD")
  })

  it('documents Relay default and the explicit rollback switch', async () => {
    const readme = await readFile(resolve(import.meta.dirname, '../README.md'), 'utf8')
    expect(readme).toContain('wss://office.8-216-134-194.sslip.io/office-relay')
    expect(readme).toContain('VITE_WISWORK_OFFICE_TRANSPORT=loopback')
    expect(readme).not.toContain('WISWORK_OFFICE_ALLOWED_ORIGIN')
  })
})
