import { readFileSync } from 'node:fs'
import { resolve } from 'node:path'
import { describe, expect, it } from 'vitest'

const root = resolve(__dirname, '../../..')
const read = (path: string) => readFileSync(resolve(root, path), 'utf8')

describe('application theme UI', () => {
  it('exposes a visible two-value theme control on Home', () => {
    const home = read('apps/shell/src/renderer/src/Home.tsx')
    expect(home).toContain('theme-switch')
    expect(home).toContain("setTheme('light')")
    expect(home).toContain("setTheme('dark')")
  })

  it.each(['docs', 'sheets', 'slides', 'pdf', 'latex'])('%s has dark chrome styles', (app) => {
    const css = read(`apps/${app}/src/renderer/styles.css`)
    expect(css).toContain('@media (prefers-color-scheme: dark)')
    expect(css).toContain('color-scheme: dark')
  })

  it('keeps authored document surfaces out of application dark-theme recoloring', () => {
    expect(read('apps/docs/src/renderer/styles.css')).toContain('--document-surface: #ffffff')
    expect(read('apps/slides/src/renderer/styles.css')).toContain(
      '--authored-slide-surface: #ffffff',
    )
    expect(read('apps/pdf/src/renderer/styles.css')).toContain('--authored-page-surface: #ffffff')
  })
})
