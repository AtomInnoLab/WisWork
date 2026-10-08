import { expect, it, vi } from 'vitest'
import { startTeamAuthPage } from '../src/team-auth-start.js'
import { completeTeamAuthCallbackPage } from '../src/team-auth-callback.js'
const config = {
  clientId: 'registered-client',
  authorizationEndpoint: 'https://login.example/authorize',
  redirectUri: 'https://addin.example/team-auth-callback.html',
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
it('redirects only the controlled same-origin bootstrap to exact configured authorization', () => {
  const replace = vi.fn()
  startTeamAuthPage(
    { origin: 'https://addin.example', hash: '#' + encodeURIComponent(authorizationUrl), replace },
    config,
  )
  expect(replace).toHaveBeenCalledWith(authorizationUrl)
  for (const hash of [
    '#' + encodeURIComponent(authorizationUrl.replace('login.example', 'evil.example')),
    '#' + encodeURIComponent(authorizationUrl + '&redirect_uri=https://evil.example'),
    '#%invalid',
  ]) {
    expect(() =>
      startTeamAuthPage({ origin: 'https://addin.example', hash, replace }, config),
    ).toThrow('team_auth_dialog_invalid')
  }
  expect(replace).toHaveBeenCalledTimes(1)
})
it('does not navigate when real configuration is missing', () => {
  const replace = vi.fn()
  expect(() =>
    startTeamAuthPage(
      {
        origin: 'https://addin.example',
        hash: '#' + encodeURIComponent(authorizationUrl),
        replace,
      },
      undefined,
    ),
  ).toThrow()
  expect(replace).not.toHaveBeenCalled()
})
it('scrubs callback query and fragment before awaiting Office readiness and sends only to exact parent origin', async () => {
  let ready!: () => void
  const events: string[] = []
  const runtime = {
    onReady: () => new Promise<void>((r) => (ready = r)),
    isDialogOriginSupported: () => true,
    messageParent: vi.fn((_message: string, _options: unknown) => events.push('message')),
  }
  const url =
    config.redirectUri +
    '?code=private-code&state=' +
    'a'.repeat(43) +
    '&iss=https%3A%2F%2Flogin.example'
  const scrub = vi.fn((_state: unknown, _unused: string, value: string) => events.push(value))
  const pending = completeTeamAuthCallbackPage({ href: url, replaceState: scrub }, config, runtime)
  expect(events).toEqual(['/team-auth-callback.html'])
  expect(runtime.messageParent).not.toHaveBeenCalled()
  ready()
  await pending
  expect(runtime.messageParent).toHaveBeenCalledWith(
    JSON.stringify({ type: 'team.auth.callback', url }),
    { targetOrigin: 'https://addin.example' },
  )
  expect(events).toEqual(['/team-auth-callback.html', 'message'])
})
it('scrubs secrets even when configuration, origin support or callback identity is unavailable', async () => {
  for (const scenario of ['unconfigured', 'foreign', 'unsupported', 'fragment']) {
    const url =
      (scenario === 'foreign'
        ? config.redirectUri.replace('addin.example', 'evil.example')
        : config.redirectUri) +
      '?code=secret&state=' +
      'a'.repeat(43) +
      (scenario === 'fragment' ? '#access_token=secret' : '')
    const replaceState = vi.fn(),
      runtime = {
        onReady: async () => {},
        isDialogOriginSupported: () => scenario !== 'unsupported',
        messageParent: vi.fn(),
      }
    await expect(
      completeTeamAuthCallbackPage(
        { href: url, replaceState },
        scenario === 'unconfigured' ? undefined : config,
        runtime,
      ),
    ).rejects.toThrow()
    expect(replaceState).toHaveBeenCalledOnce()
    expect(runtime.messageParent).not.toHaveBeenCalled()
  }
})
