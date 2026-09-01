export const ALLOWED_ENHANCED_CAPABILITIES = [
  'semantic-read',
  'transaction-proposal',
  'bounded-render-facts',
  'raw-office-proposal',
] as const

export const DENIED_ENHANCED_CAPABILITIES = [
  'shell',
  'arbitrary-filesystem',
  'git',
  'browser-control',
  'free-network',
  'direct-document-write',
] as const

export type EnhancedCapability = (typeof ALLOWED_ENHANCED_CAPABILITIES)[number]
export type CapabilityDeclaration = Readonly<{ capabilities: readonly EnhancedCapability[] }>

export const SAFE_RUNTIME_ERROR_CODES = [
  'runtime_unavailable',
  'runtime_blocked',
  'runtime_protocol_error',
  'runtime_input_rejected',
] as const
export type SafeRuntimeError = Readonly<{ code: (typeof SAFE_RUNTIME_ERROR_CODES)[number] }>
