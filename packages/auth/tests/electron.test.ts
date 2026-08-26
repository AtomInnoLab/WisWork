import { readFileSync } from 'node:fs'
import { describe, expect, it, vi } from 'vitest'
import {
  createEncryptedSessionStore,
  extractCallbackUrl,
  registerAuthProtocolRouting,
} from '../src/index'

describe('Electron adapter', () => {
  it('fails closed when safeStorage is unavailable', async () => {
    const writeFile = vi.fn()
    const store = createEncryptedSessionStore({
      path: '/tmp/session.enc',
      safeStorage: {
        isEncryptionAvailable: () => false,
        encryptString: vi.fn(),
        decryptString: vi.fn(),
      },
      readFile: vi.fn(),
      writeFile,
      unlink: vi.fn(),
      rename: vi.fn(),
      mkdir: vi.fn(),
    })
    await expect(
      store.save({ accessToken: 'secret', refreshToken: 'refresh', userId: 'u' }),
    ).rejects.toMatchObject({ code: 'secure_storage_unavailable' })
    expect(writeFile).not.toHaveBeenCalled()
  })

  it.each(['basic_text', '', 'unknown_backend'])(
    'fails closed for Linux unsafe backend %j',
    async (backend) => {
      const writeFile = vi.fn()
      const store = createEncryptedSessionStore({
        path: '/tmp/session.enc',
        platform: 'linux',
        safeStorage: {
          isEncryptionAvailable: () => true,
          getSelectedStorageBackend: () => backend,
          encryptString: vi.fn(),
          decryptString: vi.fn(),
        },
        readFile: vi.fn(),
        writeFile,
        unlink: vi.fn(),
        rename: vi.fn(),
        mkdir: vi.fn(),
      })
      await expect(
        store.save({ accessToken: 'secret', refreshToken: 'refresh', userId: 'u' }),
      ).rejects.toMatchObject({ code: 'secure_storage_unavailable' })
      expect(writeFile).not.toHaveBeenCalled()
    },
  )

  it('extracts the exact callback from macOS open-url and Windows/Linux argv', () => {
    const callback = 'wiswork://oauth/callback?code=x&state=y'
    expect(extractCallbackUrl(callback)).toBe(callback)
    expect(extractCallbackUrl(['wiswork.exe', '--flag', callback])).toBe(callback)
    expect(extractCallbackUrl(['wiswork://oauth/other?code=x', callback])).toBe(callback)
    expect(extractCallbackUrl(['wiswork.exe', 'https://example.com'])).toBeNull()
  })
})

describe('protocol lifecycle routing', () => {
  it('registers wiswork and routes open-url, initial argv, and second-instance callbacks', async () => {
    let openUrlHandler: ((event: { preventDefault(): void }, url: string) => void) | undefined
    let secondInstanceHandler: ((argv: readonly string[]) => void) | undefined
    const consume = vi.fn(async () => undefined)
    const preventDefault = vi.fn()
    const result = registerAuthProtocolRouting({
      registerProtocolClient: vi.fn(() => true),
      onOpenUrl: (handler) => {
        openUrlHandler = handler
      },
      onSecondInstance: (handler) => {
        secondInstanceHandler = handler
      },
      initialArgv: ['wiswork', 'wiswork://oauth/callback?code=initial&state=s'],
      consume,
    })
    expect(result.protocolRegistered).toBe(true)
    openUrlHandler?.({ preventDefault }, 'https://example.com')
    openUrlHandler?.({ preventDefault }, 'wiswork://oauth/callback?code=mac&state=s')
    secondInstanceHandler?.(['wiswork', 'wiswork://oauth/callback?code=second&state=s'])
    secondInstanceHandler?.(['wiswork', 'wiswork://oauth/callback?code=second&state=s'])
    await vi.waitFor(() => expect(consume).toHaveBeenCalledTimes(4))
    expect(preventDefault).toHaveBeenCalledTimes(1)
  })
})

describe('packaging integration', () => {
  const readRepo = (relative: string) =>
    readFileSync(new URL(`../../../${relative}`, import.meta.url), 'utf8')

  it('bundles auth source into every editor main process', () => {
    for (const appName of ['docs', 'sheets', 'slides']) {
      expect(readRepo(`apps/${appName}/electron.vite.config.ts`)).toContain('@wiswork/auth')
    }
  })

  it('guards every auth IPC handler and exposes no legacy withEmail payload', () => {
    const sources = [
      readRepo('apps/docs/src/main/docs-main.ts'),
      readRepo('apps/sheets/src/main/sheets-main.ts'),
      readRepo('apps/slides/src/main/ai-ipc.ts'),
    ]
    for (const source of sources) {
      expect(source).toContain('assertAuthIpc(event, args)')
      expect(source.match(/assertAuthIpc\(event, args\)/g)).toHaveLength(3)
      expect(source).not.toContain('withEmail')
    }
    const shell = readRepo('apps/shell/src/main/index.ts')
    expect(shell.match(/assertHomeAuthIpc\(event, args\)/g)).toHaveLength(6)
  })

  it('declares the wiswork protocol only in the packaged Shell', () => {
    expect(readRepo('apps/shell/electron-builder.cjs')).toMatch(
      /protocols[\s\S]*schemes:\s*\['wiswork'\]/,
    )
    for (const appName of ['docs', 'slides']) {
      const manifest = JSON.parse(readRepo('apps/' + appName + '/package.json')) as {
        build?: { protocols?: unknown }
      }
      expect(manifest.build?.protocols).toBeUndefined()
    }
    const standaloneSources = [
      [readRepo('apps/docs/src/main/docs-main.ts'), readRepo('apps/docs/src/main/docs-main.ts')],
      [
        readRepo('apps/sheets/src/main/sheets-main.ts'),
        readRepo('apps/sheets/src/main/sheets-main.ts'),
      ],
      [readRepo('apps/slides/src/main/slides-main.ts'), readRepo('apps/slides/src/main/ai-ipc.ts')],
    ]
    for (const [mainSource, authIpcSource] of standaloneSources) {
      expect(mainSource).not.toContain('setAsDefaultProtocolClient')
      expect(mainSource).not.toContain('initializeElectronAuthRuntime')
      expect(authIpcSource).toContain('auth_unavailable_in_standalone')
    }
  })
})
