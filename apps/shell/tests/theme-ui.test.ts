import { readFileSync } from 'node:fs'
import { resolve } from 'node:path'
import { describe, expect, it } from 'vitest'

const root = resolve(__dirname, '../../..')
const read = (path: string) => readFileSync(resolve(root, path), 'utf8')

describe('application theme UI', () => {
  it('exposes the two-value theme control inside the account flyout', () => {
    const home = read('apps/shell/src/renderer/src/Home.tsx')
    expect(home).toContain('theme-switch')
    expect(home).toContain("setTheme('light')")
    expect(home).toContain("setTheme('dark')")
    const menu = home.indexOf('className="account-menu"')
    const theme = home.indexOf('<ThemeSwitch', menu)
    const accountEntryEnd = home.indexOf('// ── Main component', menu)
    expect(menu).toBeGreaterThan(-1)
    expect(theme).toBeGreaterThan(menu)
    expect(theme).toBeLessThan(accountEntryEnd)
    expect(home.slice(accountEntryEnd)).not.toContain('<ThemeSwitch')
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

  it('does not let OS dark mode override an explicit Shell theme', () => {
    const css = read('apps/shell/src/renderer/src/home.css')
    expect(css).toContain(":root:not([data-theme='light']):not([data-theme='dark'])")
  })
})
