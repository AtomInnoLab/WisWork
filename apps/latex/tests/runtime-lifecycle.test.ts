import { readFile } from 'node:fs/promises'
import { describe, expect, it } from 'vitest'
import { configureLatexRuntime, resetLatexRuntimeForTests } from '../src/main/latex-main.js'
import { assertStandaloneDevelopment } from '../src/main/standalone-guard.js'

describe('LaTeX runtime lifecycle', () => {
  it('rejects runtime reconfiguration and permits an explicit test reset', () => {
    const config = {
      preloadPath: '/fixed/preload.mjs',
      rendererFile: '/fixed/index.html',
      tectonicPath: '/fixed/tectonic',
      userDataPath: '/fixed/data',
    }
    configureLatexRuntime(config)
    expect(() => configureLatexRuntime(config)).toThrow(/already configured/i)
    resetLatexRuntimeForTests()
    expect(() => configureLatexRuntime(config)).not.toThrow()
    resetLatexRuntimeForTests()
  })

  it('hard rejects packaged standalone mode and uses the emitted preload extension', async () => {
    expect(() => assertStandaloneDevelopment(true)).toThrow(/packaged/i)
    expect(() => assertStandaloneDevelopment(false)).not.toThrow()
    expect(await readFile(new URL('../src/main/standalone.ts', import.meta.url), 'utf8')).toContain(
      '../preload/index.mjs',
    )
  })
})
