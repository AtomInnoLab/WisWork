import { contextBridge, ipcRenderer } from 'electron'
import type { IpcRendererEvent } from 'electron'
import type {
  AccountLoginEvent,
  AccountStatus,
  OfficePairingRequest,
  OfficeBridgeStatus,
  OfficeRelayStatus,
  HomeApi,
  LatexRecentProjectEntry,
  RecentEntry,
  RecentPage,
  RenameResult,
  ProjectHomeApi,
  ProjectSummaryEntry,
  TimelineEntryItem,
  UiLanguage,
  AppTheme,
} from '../shared/home-api'
import { HOME_CHANNELS, OFFICE_PAIRING_CHANNELS, PROJECT_CHANNELS } from '../shared/home-api'
import type { TabsApi, TabSummary } from '../shared/tabs-api'
import { TABS_CHANNELS } from '../shared/tabs-api'

const UI_LANGUAGES: readonly UiLanguage[] = [
  'zh',
  'en',
  'ja',
  'ko',
  'fr',
  'de',
  'es',
  'th',
  'id',
  'ru',
  'ar',
  'pt',
  'it',
  'pl',
  'nl',
  'ms',
  'he',
  'hi',
  'zh-TW',
]

function isUiLanguage(value: unknown): value is UiLanguage {
  return UI_LANGUAGES.includes(value as UiLanguage)
}

function isAppTheme(value: unknown): value is AppTheme {
  return value === 'light' || value === 'dark'
}

const EMPTY_PAGE: RecentPage = { entries: [], total: 0, totalAll: 0 }

function asRecentPage(result: unknown): RecentPage {
  if (result && typeof result === 'object' && Array.isArray((result as RecentPage).entries)) {
    return result as RecentPage
  }
  return EMPTY_PAGE
}

const homeApi: HomeApi = {
  async latexRecents() {
    const result: unknown = await ipcRenderer.invoke(HOME_CHANNELS.latexRecents)
    return Array.isArray(result) ? (result as LatexRecentProjectEntry[]) : []
  },
  async newLatexProject() {
    const result: unknown = await ipcRenderer.invoke(HOME_CHANNELS.newLatexProject)
    return (result as LatexRecentProjectEntry | null) ?? null
  },
  async importLatexProject() {
    const result: unknown = await ipcRenderer.invoke(HOME_CHANNELS.importLatexProject)
    return (result as LatexRecentProjectEntry | null) ?? null
  },
  async openLatexProject(path) {
    if (typeof path !== 'string' || !path) throw new Error('Invalid path.')
    await ipcRenderer.invoke(HOME_CHANNELS.openLatexProject, path)
  },
  async recents(query) {
    return asRecentPage(await ipcRenderer.invoke(HOME_CHANNELS.recents, query))
  },
  async starred(query) {
    return asRecentPage(await ipcRenderer.invoke(HOME_CHANNELS.starred, query))
  },
  async statPaths(paths) {
    const result: unknown = await ipcRenderer.invoke(HOME_CHANNELS.statPaths, paths)
    return Array.isArray(result) ? (result as RecentEntry[]) : []
  },
  async toggleStar(path) {
    if (typeof path !== 'string' || !path) throw new Error('Invalid path.')
    await ipcRenderer.invoke(HOME_CHANNELS.toggleStar, path)
  },
  async openPath(path) {
    if (typeof path !== 'string' || !path) throw new Error('Invalid path.')
    await ipcRenderer.invoke(HOME_CHANNELS.openPath, path)
  },
  async browse() {
    await ipcRenderer.invoke(HOME_CHANNELS.browse)
  },
  async newDoc(opts) {
    await ipcRenderer.invoke(HOME_CHANNELS.newDoc, opts)
  },
  async newSheet(opts) {
    await ipcRenderer.invoke(HOME_CHANNELS.newSheet, opts)
  },
  async newSlide(opts) {
    await ipcRenderer.invoke(HOME_CHANNELS.newSlide, opts)
  },
  async newMarkdown(opts) {
    await ipcRenderer.invoke(HOME_CHANNELS.newMarkdown, opts)
  },
  async removeRecent(paths) {
    await ipcRenderer.invoke(HOME_CHANNELS.removeRecent, paths)
  },
  async revealPath(path) {
    if (typeof path !== 'string' || !path) throw new Error('Invalid path.')
    await ipcRenderer.invoke(HOME_CHANNELS.revealPath, path)
  },
  async renameFile(path, newName) {
    if (typeof path !== 'string' || !path) throw new Error('Invalid path.')
    const result: unknown = await ipcRenderer.invoke(HOME_CHANNELS.renameFile, path, newName)
    return (result ?? { ok: false, error: 'Rename failed' }) as RenameResult
  },
  async duplicateFile(path) {
    if (typeof path !== 'string' || !path) throw new Error('Invalid path.')
    await ipcRenderer.invoke(HOME_CHANNELS.duplicateFile, path)
  },
  async deleteFiles(paths) {
    await ipcRenderer.invoke(HOME_CHANNELS.deleteFiles, paths)
  },
  async openTrash() {
    await ipcRenderer.invoke(HOME_CHANNELS.openTrash)
  },
  async getLanguage() {
    const result: unknown = await ipcRenderer.invoke(HOME_CHANNELS.getLanguage)
    return isUiLanguage(result) ? result : 'zh'
  },
  async setLanguage(lang) {
    if (!isUiLanguage(lang)) throw new Error('Invalid language.')
    await ipcRenderer.invoke(HOME_CHANNELS.setLanguage, lang)
  },
  async getTheme() {
    const result: unknown = await ipcRenderer.invoke(HOME_CHANNELS.getTheme)
    return isAppTheme(result) ? result : 'light'
  },
  async setTheme(theme) {
    if (!isAppTheme(theme)) throw new Error('Invalid theme.')
    await ipcRenderer.invoke(HOME_CHANNELS.setTheme, theme)
  },
  onThemeChanged(handler) {
    const listener = (_event: IpcRendererEvent, theme: unknown) => {
      if (isAppTheme(theme)) handler(theme)
    }
    ipcRenderer.on(HOME_CHANNELS.themeChanged, listener)
    return () => ipcRenderer.removeListener(HOME_CHANNELS.themeChanged, listener)
  },
  async getUpdateChannel() {
    const result: unknown = await ipcRenderer.invoke(HOME_CHANNELS.getUpdateChannel)
    return result === 'beta' ? 'beta' : 'stable'
  },
  async setUpdateChannel(channel) {
    // validated inline: a runtime import from ../shared/update-api would be
    // shared with the update.ts preload entry and get split into a chunk,
    // which sandboxed preload scripts cannot load (window.aiOffice would
    // silently disappear). Preload entries must stay single-file bundles.
    if (channel !== 'stable' && channel !== 'beta') throw new Error('Invalid update channel.')
    await ipcRenderer.invoke(HOME_CHANNELS.setUpdateChannel, channel)
  },
  async accountStatus() {
    const result: unknown = await ipcRenderer.invoke(HOME_CHANNELS.accountStatus)
    return (result ?? { loggedIn: false }) as AccountStatus
  },
  async accountLogin() {
    const result: unknown = await ipcRenderer.invoke(HOME_CHANNELS.accountLogin)
    return result === true
  },
  onAccountLogin(handler) {
    const listener = (_event: IpcRendererEvent, ev: AccountLoginEvent) => handler(ev)
    ipcRenderer.on(HOME_CHANNELS.accountLoginEvent, listener)
    return () => ipcRenderer.removeListener(HOME_CHANNELS.accountLoginEvent, listener)
  },
  async openLoginUrl() {
    await ipcRenderer.invoke(HOME_CHANNELS.accountLoginOpenUrl)
  },
  async accountLogout() {
    await ipcRenderer.invoke(HOME_CHANNELS.accountLogout)
  },
  onOfficePairingRequested(handler) {
    const listener = (_event: IpcRendererEvent, value: unknown) => {
      if (!value || typeof value !== 'object') return
      const pairing = value as Partial<OfficePairingRequest>
      if (
        typeof pairing.pairingId === 'string' &&
        ['Word', 'Excel', 'PowerPoint'].includes(pairing.hostLabel ?? '') &&
        typeof pairing.origin === 'string' &&
        typeof pairing.verificationCode === 'string' &&
        /^\d{6}$/.test(pairing.verificationCode)
      )
        handler(pairing as OfficePairingRequest)
    }
    ipcRenderer.on(OFFICE_PAIRING_CHANNELS.requested, listener)
    return () => ipcRenderer.removeListener(OFFICE_PAIRING_CHANNELS.requested, listener)
  },
  async listOfficePairings() {
    const result: unknown = await ipcRenderer.invoke(OFFICE_PAIRING_CHANNELS.list)
    if (!Array.isArray(result)) return []
    return result.filter((value): value is OfficePairingRequest => {
      if (!value || typeof value !== 'object') return false
      const pairing = value as Partial<OfficePairingRequest>
      return (
        typeof pairing.pairingId === 'string' &&
        ['Word', 'Excel', 'PowerPoint'].includes(pairing.hostLabel ?? '') &&
        typeof pairing.origin === 'string' &&
        typeof pairing.verificationCode === 'string' &&
        /^\d{6}$/.test(pairing.verificationCode)
      )
    })
  },
  async approveOfficePairing(pairingId) {
    if (!/^[A-Za-z0-9_-]{8,128}$/.test(pairingId)) throw new Error('Invalid pairing ID.')
    return (await ipcRenderer.invoke(OFFICE_PAIRING_CHANNELS.approve, { pairingId })) === true
  },
  async rejectOfficePairing(pairingId) {
    if (!/^[A-Za-z0-9_-]{8,128}$/.test(pairingId)) throw new Error('Invalid pairing ID.')
    return (await ipcRenderer.invoke(OFFICE_PAIRING_CHANNELS.reject, { pairingId })) === true
  },
  async officeBridgeStatus() {
    const result: unknown = await ipcRenderer.invoke(HOME_CHANNELS.officeBridgeStatus)
    const value = String(result)
    return (
      value === 'disabled' ||
      value === 'error:pool_exhausted' ||
      value === 'error:invalid_config' ||
      value === 'error:bind_failed' ||
      /^ready:\d+$/.test(value)
        ? value
        : 'error:bind_failed'
    ) as OfficeBridgeStatus
  },
  async claimOfficeRelay(code) {
    if (!/^\d{6}$/.test(code)) throw new Error('Invalid verification code.')
    await ipcRenderer.invoke(OFFICE_PAIRING_CHANNELS.relayClaim, { code })
  },
  async officeRelayStatus() {
    const result: unknown = await ipcRenderer.invoke(OFFICE_PAIRING_CHANNELS.relayStatus)
    return (typeof result === 'string' ? result : 'error:invalid_status') as OfficeRelayStatus
  },
  async getAppVersion() {
    const result: unknown = await ipcRenderer.invoke(HOME_CHANNELS.getAppVersion)
    return typeof result === 'string' ? result : ''
  },
  async onboardingSeen() {
    const result: unknown = await ipcRenderer.invoke(HOME_CHANNELS.onboardingSeen)
    return result === true
  },
  async setOnboardingSeen() {
    await ipcRenderer.invoke(HOME_CHANNELS.setOnboardingSeen)
  },
}

contextBridge.exposeInMainWorld('aiOffice', homeApi)

const projectApi: ProjectHomeApi = {
  async listProjects() {
    const result: unknown = await ipcRenderer.invoke(PROJECT_CHANNELS.list)
    return Array.isArray(result) ? (result as ProjectSummaryEntry[]) : []
  },
  async listFiles(projectId) {
    const result: unknown = await ipcRenderer.invoke(PROJECT_CHANNELS.files, { projectId })
    return Array.isArray(result)
      ? result.filter((path): path is string => typeof path === 'string')
      : []
  },
  async createProject(name) {
    const result: unknown = await ipcRenderer.invoke(PROJECT_CHANNELS.create, { name })
    return result as ProjectSummaryEntry
  },
  async renameProject(id, name) {
    await ipcRenderer.invoke(PROJECT_CHANNELS.rename, { id, name })
  },
  async deleteProject(id) {
    await ipcRenderer.invoke(PROJECT_CHANNELS.delete, { id })
  },
  async moveFile(filePath, projectId) {
    await ipcRenderer.invoke(PROJECT_CHANNELS.moveFile, { filePath, projectId })
  },
  async getTimeline(projectId, limit) {
    const result: unknown = await ipcRenderer.invoke(PROJECT_CHANNELS.timeline, {
      projectId,
      limit,
    })
    return Array.isArray(result) ? (result as TimelineEntryItem[]) : []
  },
}

contextBridge.exposeInMainWorld('aiOfficeProject', projectApi)

const tabsApi: TabsApi = {
  async list() {
    const result: unknown = await ipcRenderer.invoke(TABS_CHANNELS.list)
    return Array.isArray(result) ? (result as TabSummary[]) : []
  },
  async activate(id) {
    await ipcRenderer.invoke(TABS_CHANNELS.activate, id)
  },
  async close(id) {
    await ipcRenderer.invoke(TABS_CHANNELS.close, id)
  },
  async showMenu(x, y) {
    await ipcRenderer.invoke(TABS_CHANNELS.showMenu, x, y)
  },
  async showNewMenu(x, y) {
    await ipcRenderer.invoke(TABS_CHANNELS.showNewMenu, x, y)
  },
  async reorder(id, toIndex) {
    await ipcRenderer.invoke(TABS_CHANNELS.reorder, id, toIndex)
  },
  onChanged(handler) {
    const listener = (_event: IpcRendererEvent, tabs: TabSummary[]) => handler(tabs)
    ipcRenderer.on(TABS_CHANNELS.changed, listener)
    return () => ipcRenderer.removeListener(TABS_CHANNELS.changed, listener)
  },
  notifyChromePressed() {
    ipcRenderer.send(TABS_CHANNELS.chromePressed)
  },
}

contextBridge.exposeInMainWorld('aiOfficeTabs', tabsApi)
