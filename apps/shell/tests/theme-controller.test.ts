import { mkdtempSync, rmSync } from 'node:fs'
import { tmpdir } from 'node:os'
import { join } from 'node:path'
import { afterEach, beforeEach, describe, expect, it, vi } from 'vitest'
import { createThemeController, isAppTheme, registerThemeIpc } from '../src/main/theme-controller'
import { HOME_CHANNELS } from '../src/shared/home-api'

let dir: string
let settingsPath: string

beforeEach(() => {
  dir = mkdtempSync(join(tmpdir(), 'wiswork-theme-'))
  settingsPath = join(dir, 'app-settings.json')
})

afterEach(() => {
  rmSync(dir, { recursive: true, force: true })
})

describe('theme controller', () => {
  it('accepts exactly light and dark', () => {
    expect(isAppTheme('light')).toBe(true)
    expect(isAppTheme('dark')).toBe(true)
    expect(isAppTheme('system')).toBe(false)
    expect(isAppTheme(undefined)).toBe(false)
  })

  it('defaults to light, persists changes, and restores them on restart', () => {
    const nativeTheme = { themeSource: 'system' }
    const broadcast = vi.fn()
    const controller = createThemeController({ settingsPath, nativeTheme, broadcast })

    expect(controller.initialize()).toBe('light')
    expect(nativeTheme.themeSource).toBe('light')
    controller.set('dark')
    expect(nativeTheme.themeSource).toBe('dark')
    expect(broadcast).toHaveBeenCalledWith('dark')

    const restartedNativeTheme = { themeSource: 'system' }
    const restarted = createThemeController({
      settingsPath,
      nativeTheme: restartedNativeTheme,
      broadcast: vi.fn(),
    })
    expect(restarted.initialize()).toBe('dark')
    expect(restartedNativeTheme.themeSource).toBe('dark')
  })

  it('rejects invalid values without mutating or broadcasting', () => {
    const nativeTheme = { themeSource: 'light' }
    const broadcast = vi.fn()
    const controller = createThemeController({ settingsPath, nativeTheme, broadcast })
    controller.initialize()

    expect(() => controller.set('system')).toThrow('Invalid theme')
    expect(nativeTheme.themeSource).toBe('light')
    expect(broadcast).not.toHaveBeenCalled()
  })

  it('validates payloads and rejects untrusted IPC senders', async () => {
    const handlers = new Map<string, (event: unknown, ...args: unknown[]) => unknown>()
    const ipcMain = {
      handle: (channel: string, handler: (event: unknown, ...args: unknown[]) => unknown) => {
        handlers.set(channel, handler)
      },
    }
    const controller = createThemeController({
      settingsPath,
      nativeTheme: { themeSource: 'light' },
      broadcast: vi.fn(),
    })
    controller.initialize()
    registerThemeIpc({ ipcMain, controller, isTrustedSender: (sender) => sender === 'shell' })

    const get = handlers.get(HOME_CHANNELS.getTheme)!
    const set = handlers.get(HOME_CHANNELS.setTheme)!
    expect(get({ sender: 'shell' })).toBe('light')
    expect(() => get({ sender: 'document' })).toThrow('Untrusted IPC sender')
    expect(() => get({ sender: 'shell' }, 'extra')).toThrow('Invalid theme IPC payload')
    expect(() => set({ sender: 'shell' }, 'system')).toThrow('Invalid theme')
    expect(set({ sender: 'shell' }, 'dark')).toBe('dark')
  })
})
