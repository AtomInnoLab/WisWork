import { expect, it, vi } from 'vitest'
import { createOfficeTeamLoginDialog } from '../src/agent/team-login-dialog.js'
const config = {
  addinOrigin: 'https://addin.example',
  authorizationEndpoint: 'https://login.example/authorize',
  redirectUri: 'https://addin.example/team-auth-callback.html',
  clientId: 'registered-client',
  timeoutMs: 1000,
}
const authorizationUrl =
  config.authorizationEndpoint +
  '?' +
  new URLSearchParams({
    client_id: config.clientId,
    redirect_uri: config.redirectUri,
    response_type: 'code',
    state: 'a'.repeat(43),
    scope: 'openid',
  })
const callbackUrl =
  config.redirectUri +
  '?code=private-code&state=' +
  'a'.repeat(43) +
  '&iss=https%3A%2F%2Flogin.example'
function fixture(supported = true) {
  const handlers = new Map<string, (value: unknown) => void>(),
    close = vi.fn(),
    dialog = {
      close,
      addEventHandler: vi.fn((name: string, handler: (value: unknown) => void) =>
        handlers.set(name, handler),
      ),
    }
  let opened: (result: unknown) => void = () => {}
  const runtime = {
    context: {
      requirements: { isSetSupported: vi.fn(() => supported) },
      ui: {
        displayDialogAsync: vi.fn(
          (_url: string, _options: unknown, callback: (result: unknown) => void) => {
            opened = callback
          },
        ),
      },
    },
    EventType: { DialogMessageReceived: 'message', DialogEventReceived: 'event' },
    AsyncResultStatus: { Succeeded: 'succeeded' },
  }
  const login = createOfficeTeamLoginDialog({ ...config, runtime: runtime as any })
  return {
    login,
    runtime,
    dialog,
    close,
    handlers,
    open: () => opened({ status: 'succeeded', value: dialog }),
    respond: (value: unknown) => opened(value),
  }
}
it('opens only a same-origin start page and resolves an origin-verified callback after closing', async () => {
  const f = fixture(),
    pending = f.login(authorizationUrl)
  f.open()
  const [url, options] = f.runtime.context.ui.displayDialogAsync.mock.calls[0]!
  expect(new URL(url).origin).toBe(config.addinOrigin)
  expect(new URL(url).pathname).toBe('/team-auth-start.html')
  expect(decodeURIComponent(new URL(url).hash.slice(1))).toBe(authorizationUrl)
  expect(options).toMatchObject({ displayInIframe: false })
  f.handlers.get('message')!({
    origin: config.addinOrigin,
    message: JSON.stringify({ type: 'team.auth.callback', url: callbackUrl }),
  })
  await expect(pending).resolves.toBe(callbackUrl)
  expect(f.close).toHaveBeenCalledOnce()
})
it.each(['foreign', 'missing', 'extra', 'redirect', 'token'])(
  'rejects %s callback evidence without returning an authorization code',
  async (scenario) => {
    const f = fixture(),
      pending = f.login(authorizationUrl)
    f.open()
    const message: Record<string, unknown> = { type: 'team.auth.callback', url: callbackUrl }
    if (scenario === 'extra') message.token = 'private'
    if (scenario === 'redirect') message.url = callbackUrl.replace('addin.example', 'evil.example')
    if (scenario === 'token') message.url = callbackUrl + '&access_token=private'
    f.handlers.get('message')!({
      origin:
        scenario === 'foreign'
          ? 'https://evil.example'
          : scenario === 'missing'
            ? undefined
            : config.addinOrigin,
      message: JSON.stringify(message),
    })
    await expect(pending).rejects.toThrow('team_auth_dialog_invalid')
    expect(f.close).toHaveBeenCalledOnce()
  },
)
it('fails closed without DialogOrigin and rejects wrong authorization parameters before opening', async () => {
  const unavailable = fixture(false)
  await expect(unavailable.login(authorizationUrl)).rejects.toThrow('team_auth_dialog_unavailable')
  expect(unavailable.runtime.context.ui.displayDialogAsync).not.toHaveBeenCalled()
  for (const url of [
    authorizationUrl.replace('registered-client', 'other-client'),
    authorizationUrl.replace('login.example', 'evil.example'),
    authorizationUrl + '&client_id=duplicate',
  ]) {
    const f = fixture()
    await expect(f.login(url)).rejects.toThrow('team_auth_dialog_invalid')
    expect(f.runtime.context.ui.displayDialogAsync).not.toHaveBeenCalled()
  }
})
it('rejects duplicate calls, aborts pending opening and closes a late acquired dialog', async () => {
  const f = fixture(),
    abort = new AbortController(),
    pending = f.login(authorizationUrl, abort.signal)
  await expect(f.login(authorizationUrl)).rejects.toThrow('team_auth_dialog_busy')
  abort.abort()
  await expect(pending).rejects.toThrow('team_auth_dialog_cancelled')
  f.open()
  expect(f.close).toHaveBeenCalledOnce()
})
it('bounds opening and active lifetime, handles cancellation safely and ignores duplicate callback', async () => {
  vi.useFakeTimers()
  try {
    const f = fixture(),
      pending = f.login(authorizationUrl)
    const rejected = expect(pending).rejects.toThrow('team_auth_dialog_timeout')
    await vi.advanceTimersByTimeAsync(1001)
    await rejected
    f.open()
    expect(f.close).toHaveBeenCalledOnce()
    const current = fixture(),
      ended = current.login(authorizationUrl)
    current.open()
    current.handlers.get('event')!({ error: 12006, detail: 'secret' })
    await expect(ended).rejects.toThrow('team_auth_dialog_cancelled')
    current.open()
    expect(current.close).toHaveBeenCalledOnce()
  } finally {
    vi.useRealTimers()
  }
})
it('closes an active dialog on abort and ignores a late callback carrying a code', async () => {
  const f = fixture(),
    abort = new AbortController(),
    pending = f.login(authorizationUrl, abort.signal)
  f.open()
  abort.abort()
  await expect(pending).rejects.toThrow('team_auth_dialog_cancelled')
  expect(f.close).toHaveBeenCalledOnce()
  f.handlers.get('message')!({
    origin: config.addinOrigin,
    message: JSON.stringify({ type: 'team.auth.callback', url: callbackUrl }),
  })
  expect(f.close).toHaveBeenCalledOnce()
})
it('does not return error callbacks or a different transaction state', async () => {
  for (const url of [
    config.redirectUri + '?error=access_denied&state=' + 'a'.repeat(43),
    callbackUrl.replace('a'.repeat(43), 'c'.repeat(43)),
  ]) {
    const f = fixture(),
      pending = f.login(authorizationUrl)
    f.open()
    f.handlers.get('message')!({
      origin: config.addinOrigin,
      message: JSON.stringify({ type: 'team.auth.callback', url }),
    })
    await expect(pending).rejects.toThrow('team_auth_dialog_invalid')
    expect(f.close).toHaveBeenCalledOnce()
  }
})
it.each([undefined, {}, { status: 'succeeded', value: {} }])(
  'rejects malformed SDK callbacks without an uncaught callback error',
  async (value) => {
    const f = fixture(),
      pending = f.login(authorizationUrl)
    expect(() => f.respond(value)).not.toThrow()
    await expect(pending).rejects.toThrow('team_auth_dialog_failed')
  },
)
it('closes a returned dialog even when the SDK reports a failed opening', async () => {
  const f = fixture(),
    pending = f.login(authorizationUrl)
  f.respond({ status: 'failed', value: f.dialog, error: { message: 'private-code' } })
  await expect(pending).rejects.toThrow('team_auth_dialog_failed')
  expect(f.close).toHaveBeenCalledOnce()
})
it('budgets the JSON envelope separately while retaining the exact callback URL limit', async () => {
  const prefix = config.redirectUri + '?state=' + 'a'.repeat(43) + '&code='
  const nearLimit = prefix + 'c'.repeat(4080 - prefix.length)
  const f = fixture(),
    pending = f.login(authorizationUrl)
  f.open()
  const message = JSON.stringify({ type: 'team.auth.callback', url: nearLimit })
  expect(new TextEncoder().encode(message).length).toBeGreaterThan(4096)
  f.handlers.get('message')!({ origin: config.addinOrigin, message })
  await expect(pending).resolves.toBe(nearLimit)
  for (const message of [
    JSON.stringify({ type: 'team.auth.callback', url: nearLimit + 'c'.repeat(17) }),
    ' '.repeat(8193),
  ]) {
    const rejected = fixture(),
      request = rejected.login(authorizationUrl)
    rejected.open()
    rejected.handlers.get('message')!({ origin: config.addinOrigin, message })
    await expect(request).rejects.toThrow('team_auth_dialog_invalid')
  }
})
