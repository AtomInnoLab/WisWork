import { createHash, randomBytes, timingSafeEqual } from 'node:crypto'

const SECRET_BYTES = 32

export function createBridgeSecret(): string {
  return randomBytes(SECRET_BYTES).toString('base64url')
}

export function hasValidBearerToken(header: string | undefined, expected: string): boolean {
  const match = /^Bearer ([A-Za-z0-9_-]+)$/i.exec(header ?? '')
  const presented = match?.[1] ?? ''
  const expectedDigest = createHash('sha256').update(expected, 'utf8').digest()
  const presentedDigest = createHash('sha256').update(presented, 'utf8').digest()
  const equal = timingSafeEqual(expectedDigest, presentedDigest)
  return match !== null && equal
}
