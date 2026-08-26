import { access, readFile, readdir } from 'node:fs/promises'
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
    expect(manifest).toContain('<Version>0.3.3.0</Version>')
    expect(manifest).toContain('https://office.example/taskpane.html?v=0.3.3')
    expect(manifest).not.toContain('auth.example')
    expect(manifest).not.toContain('localhost')
    expect(manifest).not.toContain('*')
  })

  it('emits one task pane with the fixed relay policy and no legacy auth assets', async () => {
    const taskpane = await readFile(resolve(dist, 'taskpane.html'), 'utf8')
    const files = await readdir(dist, { recursive: true })
    expect(taskpane).toContain("connect-src 'self' wss://office.8-216-134-194.sslip.io")
    expect(taskpane).not.toContain('http://127.0.0.1')
    const scriptPath = taskpane.match(/src="(\/assets\/taskpane-[^"]+\.js)"/)?.[1]
    expect(scriptPath).toBeDefined()
    const script = await readFile(resolve(dist, scriptPath!.replace(/^\//, '')), 'utf8')
    expect(script).toContain('wss://office.8-216-134-194.sslip.io/office-relay')
    expect(taskpane).not.toMatch(/oauth|callback|auth\.dev|wisusage/i)
    expect(taskpane).not.toContain("'unsafe-eval'")
    expect(files).not.toContain('oauth')
    expect(files.some((file) => file.startsWith('assets/conversion-worker-'))).toBe(true)
    expect(files.some((file) => file.endsWith('.map'))).toBe(false)
  })

  it('omits a deployable manifest from an unconfigured build', async () => {
    const keys = ['VITE_WISWORK_ADDIN_ORIGIN', 'VITE_WISWORK_PC_BRIDGE_PORTS']
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
  }, 15_000)

  it('can build the retained legacy workspace with only its independent rollback flag', async () => {
    const configured = {
      VITE_WISWORK_ADDIN_ORIGIN: 'https://office.example',
      VITE_WISWORK_OFFICE_WORKSPACE: '0',
    }
    const prior = Object.fromEntries(Object.keys(configured).map((key) => [key, process.env[key]]))
    Object.assign(process.env, configured)
    try {
      await build({ configFile: resolve(appRoot, 'vite.config.ts'), logLevel: 'silent' })
      const taskpane = await readFile(resolve(dist, 'taskpane.html'), 'utf8')
      const scriptPath = taskpane.match(/src="(\/assets\/taskpane-[^"]+\.js)"/)?.[1]
      expect(scriptPath).toBeDefined()
      const script = await readFile(resolve(dist, scriptPath!.replace(/^\//, '')), 'utf8')
      expect(script).toContain('WisWork Agent')
      expect(script).not.toContain('Work with your selection')
      expect(script).not.toContain('Session files')
      expect(script).toContain('legacy-workspace')
    } finally {
      for (const [key, value] of Object.entries(prior)) {
        if (value === undefined) delete process.env[key]
        else process.env[key] = value
      }
    }
  }, 15_000)
})
