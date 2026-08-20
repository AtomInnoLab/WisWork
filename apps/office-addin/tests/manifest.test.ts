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
})
