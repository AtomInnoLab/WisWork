import { readFile } from 'node:fs/promises'
import { describe, expect, it, vi } from 'vitest'
import {
  attachLatexViewSessionForHost,
  configureLatexRuntime,
  getOpeningLatexViewCountForTests,
  resetLatexRuntimeForTests,
  teardownLatexRenderer,
} from '../src/main/latex-main.js'
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
      '../preload/index.cjs',
    )
  })

  it('revokes the project session when a detached renderer cannot be closed safely', () => {
    const destroy = vi.fn()
    teardownLatexRenderer({ id: 71, isDestroyed: () => false } as never, { destroy } as never)
    expect(destroy).toHaveBeenCalledWith(71)
  })

  it('revokes an attach that completes after the renderer was torn down', async () => {
    let finishAttach!: () => void
    const attachGate = new Promise<void>((resolve) => (finishAttach = resolve))
    const sessions = new Map<number, { dispose(): void }>()
    const dispose = vi.fn()
    const registry = {
      async attach(id: number) {
        await attachGate
        const session = { projectId: 'late-session', dispose }
        sessions.set(id, session)
        return session
      },
      destroy(id: number) {
        sessions.get(id)?.dispose()
        sessions.delete(id)
      },
    }
    const send = vi.fn()
    const contents = {
      id: 72,
      isDestroyed: () => false,
      once: vi.fn(),
      send,
    }
    attachLatexViewSessionForHost(
      contents as never,
      '/project',
      registry as never,
      Promise.resolve(),
    )
    teardownLatexRenderer(contents as never, registry as never)
    finishAttach()
    await vi.waitFor(() => expect(dispose).toHaveBeenCalledOnce())
    expect(sessions.size).toBe(0)
    expect(send).not.toHaveBeenCalledWith('latex:project:opened', expect.anything())
  })

  it('retains the tombstone when load rejects before a deferred attach succeeds', async () => {
    let finishAttach!: () => void
    const attachGate = new Promise<void>((resolve) => (finishAttach = resolve))
    const sessions = new Map<number, { dispose(): void }>()
    const dispose = vi.fn()
    const registry = {
      async attach(id: number) {
        await attachGate
        const session = { projectId: 'late-after-load-error', dispose }
        sessions.set(id, session)
        return session
      },
      destroy(id: number) {
        sessions.get(id)?.dispose()
        sessions.delete(id)
      },
    }
    const send = vi.fn()
    const contents = { id: 73, isDestroyed: () => false, once: vi.fn(), send }
    attachLatexViewSessionForHost(
      contents as never,
      '/project',
      registry as never,
      Promise.reject(new Error('renderer load failed')),
    )
    await new Promise((resolve) => setImmediate(resolve))
    teardownLatexRenderer(contents as never, registry as never)
    finishAttach()
    await vi.waitFor(() => expect(dispose).toHaveBeenCalledOnce())
    expect(sessions.size).toBe(0)
    expect(getOpeningLatexViewCountForTests()).toBe(0)
    expect(send).not.toHaveBeenCalledWith('latex:project:opened', expect.anything())
  })
})
