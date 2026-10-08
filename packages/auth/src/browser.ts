import { createAuthCore, type AuthClient, type AuthClientOptions } from './oauth-core'
import { AuthError, type AuthSession, type SessionStore } from './session'
export { DEFAULT_AUTH_CONFIG } from './config'
export type { AuthConfig } from './config'
export type { AuthClient, AuthClientOptions } from './oauth-core'
export { AuthError, publicAccountStatus } from './session'
export type {
  AuthSession,
  SessionStore,
  AccountStatus,
  AuthErrorCode,
  AuthDiagnostic,
} from './session'
export type BrowserAuthClientOptions = AuthClientOptions

export function createMemorySessionStore(): SessionStore {
  let session: AuthSession | null = null
  return {
    async load() {
      return session ? { ...session } : null
    },
    async save(value) {
      session = { ...value }
    },
    async clear() {
      session = null
    },
  }
}

export function createBrowserAuthClient(options: BrowserAuthClientOptions): AuthClient {
  if (typeof options.config?.clientId !== 'string' || !options.config.clientId.trim())
    throw new AuthError('invalid_callback')
  let redirect: URL
  try {
    redirect = new URL(options.config?.redirectUri ?? '')
    if (
      redirect.protocol !== 'https:' ||
      redirect.username ||
      redirect.password ||
      redirect.search ||
      redirect.hash
    )
      throw Error('invalid_redirect')
  } catch {
    throw new AuthError('invalid_callback')
  }
  return createAuthCore(options, {
    randomBytes: (size) => globalThis.crypto.getRandomValues(new Uint8Array(size)),
    base64Url: (bytes) =>
      btoa(String.fromCharCode(...bytes))
        .replace(/\+/g, '-')
        .replace(/\//g, '_')
        .replace(/=+$/, ''),
    equalSecret(a, b) {
      if (a.length !== b.length) return false
      let difference = 0
      for (let index = 0; index < a.length; index++)
        difference |= a.charCodeAt(index) ^ b.charCodeAt(index)
      return difference === 0
    },
    validCallback: (url) => url.origin === redirect.origin && url.pathname === redirect.pathname,
  })
}
