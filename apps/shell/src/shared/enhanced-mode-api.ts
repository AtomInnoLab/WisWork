export const ENHANCED_MODE_CHANNELS = {
  status: 'enhanced-mode:component:status',
  install: 'enhanced-mode:component:install',
  remove: 'enhanced-mode:component:remove',
  update: 'enhanced-mode:component:update',
  cancel: 'enhanced-mode:component:cancel',
  setMode: 'enhanced-mode:set-mode',
  diagnostics: 'enhanced-mode:diagnostics:list',
  selfCheck: 'enhanced-mode:diagnostics:self-check',
  enableDetailed: 'enhanced-mode:diagnostics:enable-detailed',
  copyDiagnosticId: 'enhanced-mode:diagnostics:copy-id',
  exportDiagnostics: 'enhanced-mode:diagnostics:export',
} as const

export type ProductAgentMode = 'standard' | 'enhanced'
export type EnhancedModeComponentState = 'unsupported' | 'missing' | 'ready' | 'invalid'

export interface EnhancedModeStatus {
  readonly requestedAgentRuntime: ProductAgentMode
  readonly activeAgentRuntime: ProductAgentMode
  readonly component: EnhancedModeComponentState
  readonly supported: boolean
  readonly version: string
  readonly restartRequired: boolean
  readonly lifecycleState:
    | 'not_installed'
    | 'downloading'
    | 'verifying'
    | 'installed_restart_required'
    | 'ready'
    | 'update_available'
    | 'removal_restart_required'
    | 'blocked_by_policy'
    | 'unsupported_platform'
    | 'failed_safe'
  readonly platform?: string
  readonly bytes?: number
  readonly publisher?: string
  readonly license?: string
  readonly primaryUrl?: string
  readonly fallbackUrl?: string
}

export interface EnhancedDiagnosticTaskSummary {
  readonly diagnosticId: string
  readonly host: string
  readonly startedAt: number
  readonly endedAt?: number
  readonly status: 'running' | 'succeeded' | 'failed' | 'cancelled'
  readonly failureCode?: string
}

export interface EnhancedDiagnosticsSummary {
  readonly recent: readonly EnhancedDiagnosticTaskSummary[]
  readonly detailedUntil: number | null
}

export interface EnhancedSelfCheckPublicItem {
  readonly layer: 'component' | 'authentication' | 'runtime' | 'mcp' | 'wisusage'
  readonly status: 'passed' | 'failed' | 'not_tested'
  readonly code?: string
}

export interface EnhancedSelfCheckPublicResult {
  readonly diagnosticId: string
  readonly startedAt: number
  readonly endedAt: number
  readonly status: 'passed' | 'failed'
  readonly checks: readonly EnhancedSelfCheckPublicItem[]
}

export interface EnhancedModeApi {
  status(): Promise<EnhancedModeStatus>
  install(): Promise<EnhancedModeStatus>
  remove(): Promise<EnhancedModeStatus>
  update(): Promise<EnhancedModeStatus>
  cancel(): Promise<EnhancedModeStatus>
  setMode(mode: ProductAgentMode): Promise<EnhancedModeStatus>
  diagnostics(): Promise<EnhancedDiagnosticsSummary>
  selfCheck(): Promise<EnhancedSelfCheckPublicResult>
  enableDetailed(): Promise<EnhancedDiagnosticsSummary>
  copyDiagnosticId(diagnosticId: string): Promise<void>
  exportDiagnostics(): Promise<'saved' | 'cancelled'>
}
