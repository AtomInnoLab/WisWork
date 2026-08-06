import type { PathPolicyTestHooks } from './path-policy.js'

export const DEFAULT_MAX_TEXT_BYTES = 2 * 1024 * 1024

export interface LatexProjectOptions {
  maxTextBytes?: number
  savedMainFile?: string
  /** @internal Deterministic filesystem race hooks for tests; never expose through preload. */
  pathHooks?: PathPolicyTestHooks
}

export interface SavedText {
  path: string
  sha256: string
}

export type MainFileSource = 'saved' | 'tectonic' | 'main' | 'documentclass'

export type MainFileDiscovery =
  | { kind: 'found'; path: string; source: MainFileSource }
  | { kind: 'selection-required'; candidates: string[] }
  | { kind: 'not-found'; candidates: [] }

export interface MainFileDiscoveryOptions {
  savedMainFile?: string
}
