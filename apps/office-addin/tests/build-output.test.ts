import { access, readFile, stat } from 'node:fs/promises'
import { resolve } from 'node:path'
import { build } from 'vite'
import { beforeAll, describe, expect, it } from 'vitest'

const appRoot = resolve(import.meta.dirname, '..')
const dist = resolve(appRoot, 'dist')

beforeAll(async () => {
  const configured = {
    VITE_WISWORK_AUTHORIZATION_URL: 'https://auth.example/oauth/authorize',
    VITE_WISWORK_TOKEN_URL: 'https://auth.example/oauth/token',
    VITE_WISWORK_CALLBACK_URL: 'https://office.example/oauth/callback',
    VITE_WISWORK_CLIENT_ID: 'office-public',
    VITE_WISWORK_ISSUER: 'https://auth.example',
    VITE_WISWORK_MESSAGES_URL: 'https://wisusage.dev.atominnolab.com/v1/messages',
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
  it('emits the exact extensionless OAuth callback as complete HTML', async () => {
    const exactRoute = resolve(dist, 'oauth/callback')
    const [route, htmlRoute, info] = await Promise.all([
      readFile(exactRoute, 'utf8'),
      readFile(resolve(dist, 'oauth/callback.html'), 'utf8'),
      stat(exactRoute),
    ])
    expect(info.isFile()).toBe(true)
    expect(route).toBe(htmlRoute)
    expect(route).toMatch(/^<!doctype html>/)
    expect(route).toMatch(/<meta\s+http-equiv="Content-Security-Policy"/)
  })

  it('emits the same-origin Office dialog bootstrap page', async () => {
    const start = await readFile(resolve(dist, 'oauth/dialog-start.html'), 'utf8')
    expect(start).toMatch(/^<!doctype html>/)
    expect(start).toContain('main-')
    expect(start).not.toMatch(/state=|code_challenge=/)
  })

  it('emits only configured origins in the deployment manifest', async () => {
    const manifest = await readFile(resolve(dist, 'manifest.xml'), 'utf8')
    expect(manifest).toContain('https://office.example/taskpane.html')
    expect(manifest).toContain('<AppDomain>https://auth.example</AppDomain>')
    expect(manifest).not.toContain('localhost')
    expect(manifest).not.toContain('*')
  })

  it('omits a deployable manifest from an unconfigured build', async () => {
    const keys = [
      'VITE_WISWORK_AUTHORIZATION_URL',
      'VITE_WISWORK_TOKEN_URL',
      'VITE_WISWORK_CALLBACK_URL',
      'VITE_WISWORK_CLIENT_ID',
      'VITE_WISWORK_ISSUER',
      'VITE_WISWORK_MESSAGES_URL',
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
  })
})
