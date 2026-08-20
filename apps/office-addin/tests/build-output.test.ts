import { access, readFile } from 'node:fs/promises'
import { resolve } from 'node:path'
import { build } from 'vite'
import { beforeAll, describe, expect, it } from 'vitest'

const appRoot = resolve(import.meta.dirname, '..')
const dist = resolve(appRoot, 'dist')

beforeAll(async () => {
  const configured = {
    VITE_WISWORK_ADDIN_ORIGIN: 'https://office.example',
  }
  const prior = Object.fromEntries(Object.keys(configured).map((key) => [key, process.env[key]]))
  Object.assign(process.env, configured)
  try {
    await build({ configFile: resolve(appRoot, 'vite.config.ts'), logLevel: 'silent' })
  } finally {
    for (const [key, value] of Object.entries(prior)) {
      if (value === undefined) delete process.env[key]
      else process.env[key] = value
    }
  }
})

describe('configured Office build output', () => {
  it('emits only configured origins in the deployment manifest', async () => {
    const manifest = await readFile(resolve(dist, 'manifest.xml'), 'utf8')
    expect(manifest).toContain('https://office.example/taskpane.html')
    expect(manifest).not.toContain('auth.example')
    expect(manifest).not.toContain('localhost')
    expect(manifest).not.toContain('*')
  })

  it('omits a deployable manifest from an unconfigured build', async () => {
    const keys = ['VITE_WISWORK_ADDIN_ORIGIN']
    const prior = Object.fromEntries(keys.map((key) => [key, process.env[key]]))
    for (const key of keys) process.env[key] = ''
    try {
      await build({ configFile: resolve(appRoot, 'vite.config.ts'), logLevel: 'silent' })
    } finally {
      for (const [key, value] of Object.entries(prior)) {
        if (value === undefined) delete process.env[key]
        else process.env[key] = value
      }
    }
    await expect(access(resolve(dist, 'manifest.xml'))).rejects.toThrow()
  })
})
