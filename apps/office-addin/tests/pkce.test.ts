import { describe, expect, it } from 'vitest'
import { createPkcePair, createState } from '../src/auth/pkce.js'

describe('PKCE helpers', () => {
  it('creates an RFC 7636 S256 challenge without exposing the verifier', async () => {
    const pair = await createPkcePair()

    expect(pair.verifier).toMatch(/^[A-Za-z0-9_-]{43,128}$/)
    expect(pair.challenge).toMatch(/^[A-Za-z0-9_-]{43}$/)
    expect(pair.challenge).not.toBe(pair.verifier)
    expect(pair.method).toBe('S256')
  })

  it('creates independent URL-safe state values', () => {
    const first = createState()
    const second = createState()

    expect(first).toMatch(/^[A-Za-z0-9_-]{43}$/)
    expect(second).toMatch(/^[A-Za-z0-9_-]{43}$/)
    expect(first).not.toBe(second)
  })
})
