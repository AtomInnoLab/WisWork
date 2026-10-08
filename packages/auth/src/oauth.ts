import { randomBytes as nodeRandomBytes, timingSafeEqual } from 'node:crypto'
import { createAuthCore, type AuthClient, type AuthClientOptions } from './oauth-core'
export type { AuthClient, AuthClientOptions } from './oauth-core'

export function createAuthClient(options: AuthClientOptions): AuthClient {
  return createAuthCore(options, {
    randomBytes: (size) => nodeRandomBytes(size),
    base64Url: (bytes) => Buffer.from(bytes).toString('base64url'),
    equalSecret(a, b) {
      const left = Buffer.from(a),
        right = Buffer.from(b)
      return left.length === right.length && timingSafeEqual(left, right)
    },
    validCallback: (url) =>
      url.protocol === 'wiswork:' &&
      url.hostname === 'oauth' &&
      url.pathname === '/callback' &&
      !url.port,
  })
}
