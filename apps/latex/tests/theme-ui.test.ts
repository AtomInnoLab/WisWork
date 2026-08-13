import { readFileSync } from 'node:fs'
import { resolve } from 'node:path'
import { describe, expect, it } from 'vitest'

const root = resolve(__dirname, '../../..')
const read = (path: string) => readFileSync(resolve(root, path), 'utf8')

describe('LaTeX theme integration', () => {
  it('loads shared tokens and applies shell theme broadcasts', () => {
    const main = read('apps/latex/src/renderer/main.tsx')
    const preload = read('apps/latex/src/preload/index.ts')
    expect(main).toContain("import '@wiswork/ui/tokens.css'")
    expect(main).toContain('window.latexApi.getTheme()')
    expect(main).toContain('window.latexApi.onThemeChanged(applyTheme)')
    expect(preload).toContain("ipcRenderer.invoke('app:get-theme')")
    expect(preload).toContain("ipcRenderer.on('app:theme-changed'")
  })

  it('uses theme tokens for all application chrome while preserving PDF paper', () => {
    const css = read('apps/latex/src/renderer/styles.css')
    expect(css).toContain("[data-theme='dark']")
    expect(css).toContain('background: var(--latex-editor-bg)')
    expect(css).toContain('background: var(--color-bg-page)')
    expect(css).toContain('--latex-pdf-paper: #ffffff')
    expect(css).toMatch(/\.readonly-pdf-page\s*\{[\s\S]*?background: var\(--latex-pdf-paper\)/)
  })
})
