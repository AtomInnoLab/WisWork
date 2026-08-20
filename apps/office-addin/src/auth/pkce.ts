const URL_SAFE_ALPHABET = 'ABCDEFGHIJKLMNOPQRSTUVWXYZabcdefghijklmnopqrstuvwxyz0123456789-_'

function randomUrlSafe(length: number): string {
  const bytes = new Uint8Array(length)
  crypto.getRandomValues(bytes)
  return Array.from(bytes, (byte) => URL_SAFE_ALPHABET[byte & 63]).join('')
}

function base64Url(bytes: Uint8Array): string {
  let binary = ''
  for (const byte of bytes) binary += String.fromCharCode(byte)
  return btoa(binary).replaceAll('+', '-').replaceAll('/', '_').replace(/=+$/, '')
}

export interface PkcePair {
  verifier: string
  challenge: string
  method: 'S256'
}

export function createState(): string {
  return randomUrlSafe(43)
}

export async function createPkcePair(): Promise<PkcePair> {
  const verifier = randomUrlSafe(64)
  const digest = await crypto.subtle.digest('SHA-256', new TextEncoder().encode(verifier))
  return {
    verifier,
    challenge: base64Url(new Uint8Array(digest)),
    method: 'S256',
  }
}
