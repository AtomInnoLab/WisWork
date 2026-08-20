import type { UpdateChannel } from './update-api'

/** UI language; kept self-contained here (mirrors Lang in @wiswork/i18n) */
export type UiLanguage =
  | 'zh'
  | 'en'
  | 'ja'
  | 'ko'
  | 'fr'
  | 'de'
  | 'es'
  | 'th'
  | 'id'
  | 'ru'
  | 'ar'
  | 'pt'
  | 'it'
  | 'pl'
  | 'nl'
  | 'ms'
  | 'he'
  | 'hi'
  | 'zh-TW'

export type AppTheme = 'light' | 'dark'

/** a recent file entry shown on the home screen; type derives from the extension */
export interface RecentEntry {
  path: string
  name: string
  /** lowercased extension without the dot ('docx' | 'xlsx' | 'pptx') */
  ext: string
  /** last-modified time, ms since epoch */
  mtimeMs: number
  /** file size in bytes */
  sizeBytes: number
  /** whether the user starred this file */
  starred: boolean
}

/** paged query for the home file lists */
export interface RecentQuery {
  /** number of entries to skip (default 0) */
  offset?: number
  /** page size; 0 returns no entries but still reports totals (default 50) */
  limit?: number
  /** restrict to one extension ('docx' | 'xlsx' | 'pptx'); omit for all */
  ext?: string
}

export interface RecentPage {
  entries: RecentEntry[]
  /** total matching the query's ext filter */
  total: number
  /** total ignoring the ext filter (for the sidebar counters) */
  totalAll: number
}

export interface LatexRecentProjectEntry {
  path: string
  name: string
  lastOpenedAt: number
}

export interface HomeApi {
  latexRecents(): Promise<LatexRecentProjectEntry[]>
  newLatexProject(): Promise<LatexRecentProjectEntry | null>
  importLatexProject(): Promise<LatexRecentProjectEntry | null>
  openLatexProject(path: string): Promise<void>
  /** unified recents across document types, newest first (paged) */
  recents(query?: RecentQuery): Promise<RecentPage>
  /** starred files (independent of the recent list), newest first (paged) */
  starred(query?: RecentQuery): Promise<RecentPage>
  /** stat a specific set of paths (project view); missing files are skipped */
  statPaths(paths: string[]): Promise<RecentEntry[]>
  /** star / unstar a file */
  toggleStar(path: string): Promise<void>
  /** open an existing file, routing to the right module by extension */
  openPath(path: string): Promise<void>
  /** file picker accepting every supported extension, then routes */
  browse(): Promise<void>
  /** open a docs window at its start screen */
  newDoc(opts?: { projectId?: string }): Promise<void>
  /** open a sheets window */
  newSheet(opts?: { projectId?: string }): Promise<void>
  /** open a slides tab at its start screen (open-a-pptx) */
  newSlide(opts?: { projectId?: string }): Promise<void>
  /** open a blank markdown editor tab */
  newMarkdown(opts?: { projectId?: string }): Promise<void>
  /** drop entries from the recent list (does not touch the files) */
  removeRecent(paths: string[]): Promise<void>
  /** reveal the file in Finder / Explorer */
  revealPath(path: string): Promise<void>
  /** rename the file on disk (same directory) and update the recent list */
  renameFile(path: string, newName: string): Promise<RenameResult>
  /** copy the file next to itself (localized "copy" suffix before .ext) and record it as recent */
  duplicateFile(path: string): Promise<void>
  /** move files to the trash and drop them from the recent list */
  deleteFiles(paths: string[]): Promise<void>
  /** open the OS trash, where deleted files can be restored */
  openTrash(): Promise<void>
  /** current UI language (persisted in userData/app-settings.json) */
  getLanguage(): Promise<UiLanguage>
  /** switch + persist the UI language; main rebuilds its menus to match */
  setLanguage(lang: UiLanguage): Promise<void>
  /** current update channel (persisted in userData/app-settings.json; default 'stable') */
  getUpdateChannel(): Promise<UpdateChannel>
  /** switch + persist the update channel; triggers an immediate update check */
  setUpdateChannel(channel: UpdateChannel): Promise<void>
  /** Current application theme, persisted by the shell. */
  getTheme(): Promise<AppTheme>
  /** Apply and persist the application theme. */
  setTheme(theme: AppTheme): Promise<void>
  /** Observe theme changes initiated by this or another shell renderer. */
  onThemeChanged(handler: (theme: AppTheme) => void): () => void
  /** Non-sensitive WisWork account status. */
  accountStatus(): Promise<AccountStatus>
  /** Start WisWork login in the system browser. */
  accountLogin(): Promise<boolean>
  /** Subscribe to stable, non-sensitive login progress. */
  onAccountLogin(handler: (ev: AccountLoginEvent) => void): () => void
  /** Re-open the main-process-owned pending authorization URL. */
  openLoginUrl(): Promise<void>
  /** Clear the encrypted WisWork session. */
  accountLogout(): Promise<void>
  /** Observe explicit Office connection requests awaiting approval on this PC. */
  onOfficePairingRequested(handler: (pairing: OfficePairingRequest) => void): () => void
  listOfficePairings(): Promise<OfficePairingRequest[]>
  approveOfficePairing(pairingId: string): Promise<boolean>
  rejectOfficePairing(pairingId: string): Promise<boolean>
  officeBridgeStatus(): Promise<OfficeBridgeStatus>
  getAppVersion(): Promise<string>
  onboardingSeen(): Promise<boolean>
  setOnboardingSeen(): Promise<void>
}

export interface AccountStatus {
  loggedIn: boolean
  email?: string
  userId?: string
}

/** Stable, non-sensitive WisWork login progress. */
export interface AccountLoginEvent {
  phase: 'launched' | 'success' | 'error'
  error?: string
  diagnostic?: {
    stage: 'callback_exchange' | 'refresh'
    httpStatus?: number
  }
}

export interface OfficePairingRequest {
  pairingId: string
  hostLabel: 'Word' | 'Excel' | 'PowerPoint'
  origin: string
  verificationCode: string
}

export type OfficeBridgeStatus = 'disabled' | `ready:${number}` | 'error:pool_exhausted'

export interface RenameResult {
  ok: boolean
  /** the new absolute path when ok */
  path?: string
  error?: string
}

// ── Project-related APIs (P1) ────────────────────────────────

export interface ProjectSummaryEntry {
  id: string
  name: string
  createdAt: string
  updatedAt: string
  fileCount: number
  lastActiveAt: string
  isDefault: boolean
}

export interface TimelineEntryItem {
  filePath: string
  fileName: string
  chatId: string
  ts: string
  role: 'user' | 'assistant'
  preview: string
  seq: number
}

export interface ProjectHomeApi {
  /** list all projects (with file count + last-active time) */
  listProjects(): Promise<ProjectSummaryEntry[]>
  /** list existing files currently belonging to a project */
  listFiles(projectId: string): Promise<string[]>
  /** create a project */
  createProject(name: string): Promise<ProjectSummaryEntry>
  /** rename a project */
  renameProject(id: string, name: string): Promise<void>
  /** soft-delete a project */
  deleteProject(id: string): Promise<void>
  /** move a file into the given project */
  moveFile(filePath: string, projectId: string): Promise<void>
  /** fetch the project timeline */
  getTimeline(projectId: string, limit?: number): Promise<TimelineEntryItem[]>
}

export const HOME_CHANNELS = {
  recents: 'home:recents',
  starred: 'home:starred',
  statPaths: 'home:stat-paths',
  toggleStar: 'home:toggle-star',
  openPath: 'home:open-path',
  browse: 'home:browse',
  newDoc: 'home:new-doc',
  newSheet: 'home:new-sheet',
  newSlide: 'home:new-slide',
  latexRecents: 'home:latex-recents',
  newLatexProject: 'home:new-latex-project',
  importLatexProject: 'home:import-latex-project',
  openLatexProject: 'home:open-latex-project',
  newMarkdown: 'home:new-markdown',
  removeRecent: 'home:remove-recent',
  revealPath: 'home:reveal-path',
  renameFile: 'home:rename-file',
  duplicateFile: 'home:duplicate-file',
  deleteFiles: 'home:delete-files',
  openTrash: 'home:open-trash',
  getLanguage: 'home:get-language',
  setLanguage: 'home:set-language',
  getTheme: 'home:get-theme',
  setTheme: 'home:set-theme',
  themeChanged: 'home:theme-changed',
  getUpdateChannel: 'home:get-update-channel',
  setUpdateChannel: 'home:set-update-channel',
  accountStatus: 'home:account-status',
  accountLogin: 'home:account-login',
  accountLoginEvent: 'home:account-login-event',
  accountLoginOpenUrl: 'home:account-login-open-url',
  accountLogout: 'home:account-logout',
  officeBridgeStatus: 'home:office-bridge-status',
  getAppVersion: 'home:get-app-version',
  onboardingSeen: 'home:onboarding-seen',
  setOnboardingSeen: 'home:set-onboarding-seen',
} as const

export const PROJECT_CHANNELS = {
  list: 'project:list',
  files: 'project:files',
  create: 'project:create',
  rename: 'project:rename',
  delete: 'project:delete',
  moveFile: 'project:moveFile',
  timeline: 'project:timeline',
} as const

export const OFFICE_PAIRING_CHANNELS = {
  requested: 'home:office-pairing-requested',
  list: 'home:office-pairing-list',
  approve: 'home:office-pairing-approve',
  reject: 'home:office-pairing-reject',
} as const
