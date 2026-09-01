export const ENHANCED_MODE_CHANNELS = {
  status: 'enhanced-mode:component:status',
  install: 'enhanced-mode:component:install',
  remove: 'enhanced-mode:component:remove',
  update: 'enhanced-mode:component:update',
  cancel: 'enhanced-mode:component:cancel',
  setMode: 'enhanced-mode:set-mode',
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

export interface EnhancedModeApi {
  status(): Promise<EnhancedModeStatus>
  install(): Promise<EnhancedModeStatus>
  remove(): Promise<EnhancedModeStatus>
  update(): Promise<EnhancedModeStatus>
  cancel(): Promise<EnhancedModeStatus>
  setMode(mode: ProductAgentMode): Promise<EnhancedModeStatus>
}
