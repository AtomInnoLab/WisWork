import { readFile } from 'node:fs/promises'
import { resolve } from 'node:path'
import { describe, expect, it } from 'vitest'
import {
  deploymentConfig,
  renderDeploymentManifest,
  rewriteOAuthCallbackRequest,
} from '../build-config.js'

const manifestPath = resolve(import.meta.dirname, '../public/manifest.xml')
const validEnv = {
  VITE_WISWORK_AUTHORIZATION_URL: 'https://auth.example/oauth/authorize',
  VITE_WISWORK_TOKEN_URL: 'https://auth.example/oauth/token',
  VITE_WISWORK_CALLBACK_URL: 'https://office.example/oauth/callback',
  VITE_WISWORK_CLIENT_ID: 'office-public',
  VITE_WISWORK_ISSUER: 'https://auth.example',
  VITE_WISWORK_MESSAGES_URL: 'https://wisusage.dev.atominnolab.com/v1/messages',
}

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
    expect(manifest).toContain('<AppDomain>https://auth.example</AppDomain>')
    expect(manifest).not.toContain('localhost')
    expect(manifest).not.toContain('*')
  })

  it.each([
    { ...validEnv, VITE_WISWORK_CALLBACK_URL: 'https://office.example/not-callback' },
    { ...validEnv, VITE_WISWORK_MESSAGES_URL: 'https://attacker.example/v1/messages' },
    { ...validEnv, VITE_WISWORK_AUTHORIZATION_URL: 'https://*.example/oauth' },
  ])('rejects unsafe deployment configuration', (env) => {
    expect(deploymentConfig(env)).toBeUndefined()
  })

  it('declares callback entry and deterministic exact-route handling', async () => {
    const callback = await readFile(
      resolve(import.meta.dirname, '../src/oauth/callback.html'),
      'utf8',
    )
    const viteConfig = await readFile(resolve(import.meta.dirname, '../vite.config.ts'), 'utf8')
    expect(callback).toContain('../main.tsx')
    expect(callback).not.toMatch(/access[_-]?token|refresh[_-]?token/i)
    expect(viteConfig).toContain("fileName: 'oauth/callback'")
    expect(rewriteOAuthCallbackRequest('/oauth/callback?code=secret&state=s')).toBe(
      '/oauth/callback.html?code=secret&state=s',
    )
    expect(rewriteOAuthCallbackRequest('/oauth/callback/other')).toBe('/oauth/callback/other')
    expect(viteConfig).toContain('configurePreviewServer')
    expect(viteConfig).not.toContain("'Access-Control-Allow-Origin': '*'")
  })
})
