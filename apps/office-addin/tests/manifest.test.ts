import { readFile } from 'node:fs/promises'
import { resolve } from 'node:path'
import { describe, expect, it } from 'vitest'

const manifestPath = resolve(import.meta.dirname, '../public/manifest.xml')

describe('Office Add-in manifest', () => {
  it('targets the three supported document hosts over local HTTPS', async () => {
    const manifest = await readFile(manifestPath, 'utf8')

    expect(manifest).toContain('<Host Name="Document" />')
    expect(manifest).toContain('<Host Name="Workbook" />')
    expect(manifest).toContain('<Host Name="Presentation" />')
    expect(manifest).toContain(
      '<SourceLocation DefaultValue="https://localhost:3000/taskpane.html" />',
    )
  })

  it('uses an Office-supported raster icon URL', async () => {
    const manifest = await readFile(manifestPath, 'utf8')

    expect(manifest).toContain('<IconUrl DefaultValue="https://localhost:3000/assets/icon.png" />')
    expect(manifest).not.toContain('.svg')
  })

  it('declares only the local add-in domain and includes the OAuth callback build entry', async () => {
    const manifest = await readFile(manifestPath, 'utf8')
    const callback = await readFile(
      resolve(import.meta.dirname, '../src/oauth/callback.html'),
      'utf8',
    )
    const taskpane = await readFile(resolve(import.meta.dirname, '../src/taskpane.html'), 'utf8')
    const viteConfig = await readFile(resolve(import.meta.dirname, '../vite.config.ts'), 'utf8')

    expect(manifest).toContain('<AppDomain>https://localhost:3000</AppDomain>')
    expect(callback).toContain('../main.tsx')
    expect(callback).not.toMatch(/access[_-]?token|refresh[_-]?token/i)
    expect(callback).toContain('__WISWORK_CONNECT_ORIGINS__')
    expect(taskpane).toContain('__WISWORK_CONNECT_ORIGINS__')
    expect(`${callback}\n${taskpane}`).not.toMatch(/connect-src[^;]*\*/)
    expect(viteConfig).toContain('envDir: here')
  })
})
