import { readAppSettings, writeAppSetting } from './app-settings'
import { HOME_CHANNELS, type AppTheme } from '../shared/home-api'

interface NativeThemeLike {
  themeSource: string
}

interface IpcMainLike {
  handle(
    channel: string,
    handler: (event: { sender: unknown }, ...args: unknown[]) => unknown,
  ): void
}

export function isAppTheme(value: unknown): value is AppTheme {
  return value === 'light' || value === 'dark'
}

export function createThemeController(deps: {
  settingsPath: string
  nativeTheme: NativeThemeLike
  broadcast: (theme: AppTheme) => void
}) {
  let current: AppTheme = 'light'

  return {
    initialize(): AppTheme {
      const saved = readAppSettings(deps.settingsPath).theme
      current = isAppTheme(saved) ? saved : 'light'
      deps.nativeTheme.themeSource = current
      return current
    },
    get(): AppTheme {
      return current
    },
    set(value: unknown): AppTheme {
      if (!isAppTheme(value)) throw new Error('Invalid theme.')
      if (value === current) return current
      current = value
      writeAppSetting(deps.settingsPath, 'theme', current)
      deps.nativeTheme.themeSource = current
      deps.broadcast(current)
      return current
    },
  }
}

export type ThemeController = ReturnType<typeof createThemeController>

export function registerThemeIpc(deps: {
  ipcMain: IpcMainLike
  controller: ThemeController
  isTrustedSender: (sender: unknown) => boolean
}): void {
  const assertSender = (event: { sender: unknown }, args: readonly unknown[]) => {
    if (!deps.isTrustedSender(event.sender)) throw new Error('Untrusted IPC sender.')
    if (args.length !== 0) throw new Error('Invalid theme IPC payload.')
  }

  deps.ipcMain.handle(HOME_CHANNELS.getTheme, (event, ...args) => {
    assertSender(event, args)
    return deps.controller.get()
  })
  deps.ipcMain.handle(HOME_CHANNELS.setTheme, (event, value, ...args) => {
    assertSender(event, args)
    return deps.controller.set(value)
  })
}
