import { readFileSync } from 'node:fs'
import { resolve } from 'node:path'
import { describe, expect, it } from 'vitest'

const root = resolve(__dirname, '../../..')
const read = (path: string) => readFileSync(resolve(root, path), 'utf8')

describe('fixed application appearance', () => {
  it('does not expose Light or Dark theme controls or IPC', () => {
    const home = read('apps/shell/src/renderer/src/Home.tsx')
    const preload = read('apps/shell/src/preload/index.ts')
    const api = read('apps/shell/src/shared/home-api.ts')
    expect(home).not.toContain('ThemeSwitch')
    expect(home).not.toContain("setTheme('light')")
    expect(home).not.toContain("setTheme('dark')")
    expect(preload).not.toContain('getTheme')
    expect(preload).not.toContain('setTheme')
    expect(api).not.toContain('themeChanged')
  })

  it.each(['docs', 'sheets', 'slides', 'pdf', 'latex'])(
    '%s has no application dark-mode override',
    (app) => {
      const css = read(`apps/${app}/src/renderer/styles.css`)
      expect(css).not.toContain('@media (prefers-color-scheme: dark)')
    },
  )
})
