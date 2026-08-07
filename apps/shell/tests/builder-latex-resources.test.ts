import { createRequire } from 'node:module'
import { readFileSync } from 'node:fs'
import { describe, expect, it } from 'vitest'

const require = createRequire(import.meta.url)
const config = require('../electron-builder.cjs') as {
  electronVersion: string
  extraResources: Array<{ from: string; to: string }>
  mac: { extraResources: Array<{ from: string; to: string }> }
  win: { extraResources: Array<{ from: string; to: string }> }
}

describe('LaTeX packaged resources', () => {
  it('packages the module and platform executable without a bundle', () => {
    expect(config.extraResources).toContainEqual({ from: '../latex/out', to: 'modules/latex' })
    expect(config.mac.extraResources).toContainEqual({
      from: '../latex/native/tectonic',
      to: 'native/tectonic',
    })
    expect(config.win.extraResources).toContainEqual({
      from: '../latex/native/tectonic.exe',
      to: 'native/tectonic.exe',
    })
    expect(JSON.stringify(config)).not.toMatch(/bundle/i)
  })

  it('resolves the installed executable name per platform at runtime', () => {
    const source = readFileSync(new URL('../src/main/index.ts', import.meta.url), 'utf8')
    expect(source).toContain("process.platform === 'win32' ? 'tectonic.exe' : 'tectonic'")
    expect(source).toContain("join(process.resourcesPath, 'native', TECTONIC_EXE)")
  })

  it('pins the exact safe Electron version installed by the lockfile', () => {
    const installed = (require('electron/package.json') as { version: string }).version
    const lock = JSON.parse(
      readFileSync(new URL('../../../package-lock.json', import.meta.url), 'utf8'),
    ) as { packages: Record<string, { version?: string }> }
    const locked = lock.packages['node_modules/electron']?.version
    expect(config.electronVersion).toMatch(/^\d+\.\d+\.\d+$/)
    expect(config.electronVersion).toBe(installed)
    expect(config.electronVersion).toBe(locked)
  })
})
