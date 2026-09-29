import { officeTeamAuthConfig } from './agent/team-auth-config.js'
import { type OfficeTeamLoginDialogOptions } from './agent/team-login-dialog.js'
type Config = Pick<OfficeTeamLoginDialogOptions, 'redirectUri'>
export interface TeamAuthCallbackRuntime {
  onReady(): Promise<unknown>
  isDialogOriginSupported(): boolean
  messageParent(message: string, options: { targetOrigin: string }): void
}
export async function completeTeamAuthCallbackPage(
  page: { href: string; replaceState(data: unknown, unused: string, url: string): void },
  config: Config | undefined,
  runtime: TeamAuthCallbackRuntime,
): Promise<void> {
  const href = page.href
  let callback: URL
  try {
    callback = new URL(href)
  } catch {
    throw Error('team_auth_dialog_invalid')
  }
  // Remove the authorization response before any asynchronous API or messaging operation.
  page.replaceState(null, '', callback.pathname)
  if (!config || new TextEncoder().encode(href).length > 4096)
    throw Error('team_auth_dialog_invalid')
  const redirect = new URL(config.redirectUri)
  if (
    callback.protocol !== 'https:' ||
    callback.username ||
    callback.password ||
    callback.origin !== redirect.origin ||
    callback.pathname !== redirect.pathname ||
    callback.hash ||
    redirect.search ||
    redirect.hash
  )
    throw Error('team_auth_dialog_invalid')
  const keys = Array.from(callback.searchParams.keys())
  if (
    keys.some((key) => !['code', 'state', 'iss', 'error', 'error_description'].includes(key)) ||
    keys.length !== new Set(keys).size
  )
    throw Error('team_auth_dialog_invalid')
  await runtime.onReady()
  if (!runtime.isDialogOriginSupported()) throw Error('team_auth_dialog_unavailable')
  runtime.messageParent(JSON.stringify({ type: 'team.auth.callback', url: href }), {
    targetOrigin: redirect.origin,
  })
}
if (typeof window !== 'undefined') {
  const href = window.location.href
  // Scrub even when configuration or Office initialization is unavailable.
  window.history.replaceState(null, '', window.location.pathname)
  void (async () => {
    const config = officeTeamAuthConfig(import.meta.env, window.location.origin)
    if (typeof Office === 'undefined') throw Error('team_auth_dialog_unavailable')
    await completeTeamAuthCallbackPage(
      { href, replaceState: (data, unused, url) => window.history.replaceState(data, unused, url) },
      config,
      {
        onReady: () => Office.onReady(),
        isDialogOriginSupported: () =>
          Office.context.requirements.isSetSupported('DialogOrigin', '1.1'),
        messageParent: (message, options) => Office.context.ui.messageParent(message, options),
      },
    )
  })().catch(() => {
    const status = document.getElementById('team-auth-status')
    if (status) status.textContent = '团队登录回执无法核对，请关闭窗口并重新登录。'
  })
}
