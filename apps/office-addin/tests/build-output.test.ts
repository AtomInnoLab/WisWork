import { access, readFile, readdir } from 'node:fs/promises'
import { resolve } from 'node:path'
import { build } from 'vite'
import { beforeAll, describe, expect, it } from 'vitest'

const appRoot = resolve(import.meta.dirname, '..')
const dist = resolve(appRoot, 'dist')

beforeAll(async () => {
  const configured = {
    VITE_WISWORK_ADDIN_ORIGIN: 'https://office.example',
    VITE_WISWORK_TEAM_CLIENT_ID: 'registered-fixture-client',
    VITE_WISWORK_TEAM_REDIRECT_URI: 'https://office.example/team-auth-callback.html',
    VITE_WISWORK_PRESENTATION_ROLLOUT_PERCENT: '25',
    VITE_WISWORK_OFFICE_DIAGNOSTIC_SAMPLE_PERCENT: '10',
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
}, 30_000) // A production Vite build can exceed the default hook budget alongside other test workers.

describe('configured Office build output', () => {
  it('emits version metadata matching the compiled task pane', async () => {
    const metadata = JSON.parse(await readFile(resolve(dist, 'version.json'), 'utf8')) as {
      buildId: string
      presentationRolloutPercent: number
      diagnosticSamplePercent: number
      presentationMinPcProtocol: number
      presentationMinRelayProtocol: number
    }
    expect(metadata.buildId).toMatch(/^[A-Za-z0-9_.-]{3,96}$/)
    expect(metadata.presentationRolloutPercent).toBe(25)
    expect(metadata.diagnosticSamplePercent).toBe(10)
    expect(metadata.presentationMinPcProtocol).toBe(2)
    expect(metadata.presentationMinRelayProtocol).toBe(2)
    const taskpane = await readFile(resolve(dist, 'taskpane.html'), 'utf8')
    const scriptPath = taskpane.match(/src="(\/assets\/taskpane-[^"]+\.js)"/)?.[1]
    expect(scriptPath).toBeDefined()
    const script = await readFile(resolve(dist, scriptPath!.slice(1)), 'utf8')
    expect(script).toContain(metadata.buildId)
  })
  it('emits only configured origins in the deployment manifest', async () => {
    const manifest = await readFile(resolve(dist, 'manifest.xml'), 'utf8')
    expect(manifest).toContain('<Version>0.3.42.0</Version>')
    expect(manifest).toContain('https://office.example/taskpane.html?v=0.3.42')
    expect(manifest).not.toContain('auth.example')
    expect(manifest).not.toContain('localhost')
    expect(manifest).not.toContain('*')
  })

  it('emits a task pane and constrained team dialog pages without legacy auth assets', async () => {
    const taskpane = await readFile(resolve(dist, 'taskpane.html'), 'utf8')
    const files = (await readdir(dist, { recursive: true })).map((file) =>
      file.replaceAll('\\', '/'),
    )
    expect(taskpane).toContain("connect-src 'self' wss://office.8-216-134-194.sslip.io")
    expect(taskpane).toContain('https://gateway.wispaper.ai')
    for (const page of ['team-auth-start.html', 'team-auth-callback.html']) {
      const html = await readFile(resolve(dist, page), 'utf8')
      expect(html).toContain('name="referrer" content="no-referrer"')
      expect(html).toContain("connect-src 'none'")
      expect(html).toContain('https://appsforoffice.microsoft.com/lib/1/hosted/office.js')
      expect(html).not.toContain('unsafe-inline')
      expect(html).not.toContain('access_token=')
    }
    expect(taskpane).not.toContain('http://127.0.0.1')
    const scriptPath = taskpane.match(/src="(\/assets\/taskpane-[^"]+\.js)"/)?.[1]
    expect(scriptPath).toBeDefined()
    const script = await readFile(resolve(dist, scriptPath!.replace(/^\//, '')), 'utf8')
    expect(script).toContain('wss://office.8-216-134-194.sslip.io/office-relay')
    expect(taskpane).not.toMatch(/oauth|callback|auth\.dev|wisusage/i)
    expect(taskpane).not.toContain("'unsafe-eval'")
    expect(files).not.toContain('oauth')
    const conversionWorker = files.find((file) => file.startsWith('assets/conversion-worker-'))
    const pdfWorker = files.find(
      (file) => file.startsWith('assets/pdf.worker-') && file.endsWith('.js'),
    )
    expect(conversionWorker).toBeDefined()
    expect(pdfWorker).toBeDefined()
    expect(await readFile(resolve(dist, conversionWorker!), 'utf8')).toContain(
      `./${pdfWorker!.replace(/^assets\//, '')}`,
    )
    expect(files.some((file) => file.endsWith('.map'))).toBe(false)
  })

  it('omits a deployable manifest from an unconfigured build', async () => {
    const keys = [
      'VITE_WISWORK_ADDIN_ORIGIN',
      'VITE_WISWORK_PC_BRIDGE_PORTS',
      'VITE_WISWORK_TEAM_CLIENT_ID',
      'VITE_WISWORK_TEAM_REDIRECT_URI',
    ]
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
  }, 30_000)

  it('fails the build for an invalid persistent-pairing rollback flag', async () => {
    const key = 'VITE_WISWORK_OFFICE_PAIRING_RESUME'
    const prior = process.env[key]
    process.env[key] = 'false'
    try {
      await expect(
        build({ configFile: resolve(appRoot, 'vite.config.ts'), logLevel: 'silent' }),
      ).rejects.toThrow('invalid_office_pairing_resume')
    } finally {
      if (prior === undefined) delete process.env[key]
      else process.env[key] = prior
    }
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
  }, 30_000)
})
