import { describe, expect, it } from 'vitest'
import { loadRuntimeConfig } from '../src/config.js'

const validEnv = {
  VITE_WISWORK_AUTHORIZATION_URL: 'https://gateway.example/oauth/authorize',
  VITE_WISWORK_TOKEN_URL: 'https://gateway.example/oauth/token',
  VITE_WISWORK_CALLBACK_URL: 'https://localhost:3000/oauth/callback',
  VITE_WISWORK_CLIENT_ID: 'office-addin',
  VITE_WISWORK_ISSUER: 'https://gateway.example',
  VITE_WISWORK_MESSAGES_URL: 'https://gateway.example/v1/messages',
}

describe('loadRuntimeConfig', () => {
  it('returns validated, fixed browser configuration', () => {
    expect(loadRuntimeConfig(validEnv, { production: false })).toEqual({
      status: 'available',
      config: {
        authorizationUrl: validEnv.VITE_WISWORK_AUTHORIZATION_URL,
        tokenUrl: validEnv.VITE_WISWORK_TOKEN_URL,
        callbackUrl: validEnv.VITE_WISWORK_CALLBACK_URL,
        clientId: validEnv.VITE_WISWORK_CLIENT_ID,
        issuer: validEnv.VITE_WISWORK_ISSUER,
        messagesUrl: validEnv.VITE_WISWORK_MESSAGES_URL,
      },
    })
  })

  it.each([
    [{ ...validEnv, VITE_WISWORK_CLIENT_ID: '' }],
    [{ ...validEnv, VITE_WISWORK_TOKEN_URL: 'not a url' }],
    [{ ...validEnv, VITE_WISWORK_ISSUER: '' }],
    [{ ...validEnv, VITE_WISWORK_MESSAGES_URL: 'javascript:alert(1)' }],
    [{ ...validEnv, VITE_WISWORK_CALLBACK_URL: 'javascript:alert(1)' }],
  ])('fails closed for missing or invalid configuration', (env) => {
    expect(loadRuntimeConfig(env, { production: false })).toEqual({ status: 'unavailable' })
  })

  it('rejects a non-HTTPS production callback', () => {
    expect(
      loadRuntimeConfig(
        { ...validEnv, VITE_WISWORK_CALLBACK_URL: 'http://localhost:3000/oauth/callback' },
        { production: true },
      ),
    ).toEqual({ status: 'unavailable' })
  })
})
