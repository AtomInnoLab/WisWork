export const ENHANCED_MODE_CHANNELS = {
  status: 'enhanced-mode:component:status',
  install: 'enhanced-mode:component:install',
  remove: 'enhanced-mode:component:remove',
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
}

export interface EnhancedModeApi {
  status(): Promise<EnhancedModeStatus>
  install(): Promise<EnhancedModeStatus>
  remove(): Promise<EnhancedModeStatus>
  setMode(mode: ProductAgentMode): Promise<EnhancedModeStatus>
}
