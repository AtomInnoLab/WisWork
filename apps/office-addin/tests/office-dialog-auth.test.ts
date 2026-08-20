import { describe, expect, it, vi } from 'vitest'
import {
  createOfficeDialogAuth,
  authorizationFromDialogStart,
  createDialogStartUrl,
  createBrowserOfficeDialogRuntime,
  forwardOAuthCallbackToParent,
  dialogMessageOriginIsTrusted,
  officeDialogError,
  type OfficeAuthDialog,
  type OfficeDialogRuntime,
} from '../src/auth/office-dialog-auth.js'

const callbackUrl = 'https://office.example/oauth/callback'

function dialogHarness() {
  let message: ((value: string) => void) | undefined
  let error: ((code?: number) => void) | undefined
  const close = vi.fn()
  const dialog: OfficeAuthDialog = {
    close,
    onMessage(listener) {
      message = listener
      return () => (message = undefined)
    },
    onError(listener) {
      error = listener
      return () => (error = undefined)
    },
  }
  const runtime: OfficeDialogRuntime = { open: vi.fn(async () => dialog) }
  return {
    runtime,
    close,
    send: (value: string) => message?.(value),
    fail: (code?: number) => error?.(code),
  }
}

describe('Office dialog OAuth', () => {
  it('opens through a same-origin bootstrap without putting OAuth state in a request query', () => {
    const authorizationUrl =
      'https://auth.example/oidc/auth?state=secret-state&code_challenge=challenge'
    const start = new URL(createDialogStartUrl(authorizationUrl, callbackUrl))

    expect(start.origin).toBe('https://office.example')
    expect(start.pathname).toBe('/oauth/dialog-start.html')
    expect(start.search).toBe('')
    expect(authorizationFromDialogStart(start.href, 'https://auth.example/oidc/auth')).toBe(
      authorizationUrl,
    )
    expect(
      authorizationFromDialogStart(start.href, 'https://different.example/oidc/auth'),
    ).toBeUndefined()
  })

  it('opens Wispaper in an Office dialog and consumes the callback in the task pane', async () => {
    const harness = dialogHarness()
    const auth = {
      startAuthorization: vi.fn(async () => 'https://auth.example/oidc/auth?state=one'),
      consumeCallback: vi.fn(async () => undefined),
      cancelAuthorization: vi.fn(),
    }
    const signIn = createOfficeDialogAuth(harness.runtime, callbackUrl)

    const result = signIn(auth)
    await vi.waitFor(() => expect(harness.runtime.open).toHaveBeenCalledOnce())
    harness.send(
      JSON.stringify({
        type: 'wiswork_oauth_callback',
        url: `${callbackUrl}?code=one-time&state=one`,
      }),
    )
    await expect(result).resolves.toBeUndefined()

    expect(auth.consumeCallback).toHaveBeenCalledWith(`${callbackUrl}?code=one-time&state=one`)
    expect(harness.close).toHaveBeenCalledOnce()
  })

  it('rejects malformed, wrong-origin, and dialog-error messages without exchanging', async () => {
    for (const value of [
      'not-json',
      JSON.stringify({ type: 'other', url: `${callbackUrl}?code=x&state=y` }),
      JSON.stringify({
        type: 'wiswork_oauth_callback',
        url: 'https://attacker.example/oauth/callback?code=x&state=y',
      }),
    ]) {
      const harness = dialogHarness()
      const auth = {
        startAuthorization: vi.fn(async () => 'https://auth.example/oidc/auth'),
        consumeCallback: vi.fn(async () => undefined),
        cancelAuthorization: vi.fn(),
      }
      const result = createOfficeDialogAuth(harness.runtime, callbackUrl)(auth)
      await vi.waitFor(() => expect(harness.runtime.open).toHaveBeenCalledOnce())
      harness.send(value)
      await expect(result).rejects.toThrow('invalid_callback')
      expect(auth.consumeCallback).not.toHaveBeenCalled()
      expect(auth.cancelAuthorization).toHaveBeenCalledOnce()
      expect(harness.close).toHaveBeenCalledOnce()
    }

    const harness = dialogHarness()
    const auth = {
      startAuthorization: vi.fn(async () => 'https://auth.example/oidc/auth'),
      consumeCallback: vi.fn(async () => undefined),
      cancelAuthorization: vi.fn(),
    }
    const result = createOfficeDialogAuth(harness.runtime, callbackUrl)(auth)
    await vi.waitFor(() => expect(harness.runtime.open).toHaveBeenCalledOnce())
    harness.fail()
    await expect(result).rejects.toThrow('sign_in_failed')
    expect(auth.cancelAuthorization).toHaveBeenCalledOnce()
    expect(harness.close).toHaveBeenCalledOnce()
  })

  it('clears pending authorization when the dialog cannot open', async () => {
    const auth = {
      startAuthorization: vi.fn(async () => 'https://auth.example/oidc/auth'),
      consumeCallback: vi.fn(async () => undefined),
      cancelAuthorization: vi.fn(),
    }
    const runtime: OfficeDialogRuntime = {
      open: vi.fn(async () => {
        throw new Error('sign_in_failed')
      }),
    }

    await expect(createOfficeDialogAuth(runtime, callbackUrl)(auth)).rejects.toThrow(
      'sign_in_failed',
    )
    expect(auth.cancelAuthorization).toHaveBeenCalledOnce()
  })

  it('accepts an unavailable origin only in the legacy DialogOrigin compatibility path', () => {
    expect(dialogMessageOriginIsTrusted(undefined, 'https://office.example', false)).toBe(true)
    expect(dialogMessageOriginIsTrusted(undefined, 'https://office.example', true)).toBe(false)
    expect(
      dialogMessageOriginIsTrusted('https://attacker.example', 'https://office.example', false),
    ).toBe(false)
  })

  it('exposes only allowlisted Office dialog error codes', async () => {
    expect(officeDialogError(12002).message).toBe('office_dialog_12002')
    expect(officeDialogError(12006).message).toBe('office_dialog_12006')
    expect(officeDialogError(99999).message).toBe('sign_in_failed')
    expect(officeDialogError(undefined).message).toBe('sign_in_failed')

    const harness = dialogHarness()
    const auth = {
      startAuthorization: vi.fn(async () => 'https://auth.example/oidc/auth'),
      consumeCallback: vi.fn(async () => undefined),
      cancelAuthorization: vi.fn(),
    }
    const result = createOfficeDialogAuth(harness.runtime, callbackUrl)(auth)
    await vi.waitFor(() => expect(harness.runtime.open).toHaveBeenCalledOnce())
    harness.fail(12006)
    await expect(result).rejects.toThrow('office_dialog_12006')
  })

  it('forwards only an exact callback URL to the parent task pane', () => {
    const messageParent = vi.fn()
    expect(
      forwardOAuthCallbackToParent(`${callbackUrl}?code=one&state=two`, callbackUrl, messageParent),
    ).toBe(true)
    expect(messageParent).toHaveBeenCalledWith(
      JSON.stringify({
        type: 'wiswork_oauth_callback',
        url: `${callbackUrl}?code=one&state=two`,
      }),
      'https://office.example',
    )

    expect(
      forwardOAuthCallbackToParent(
        'https://attacker.example/oauth/callback?code=one&state=two',
        callbackUrl,
        messageParent,
      ),
    ).toBe(false)
  })
})
