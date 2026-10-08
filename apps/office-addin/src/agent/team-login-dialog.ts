export interface OfficeTeamLoginDialogRuntime {
  context: {
    requirements: { isSetSupported(name: string, version: string): boolean }
    ui: {
      displayDialogAsync(
        url: string,
        options: { displayInIframe: false; height: number; width: number },
        callback: (result: unknown) => void,
      ): void
    }
  }
  EventType: { DialogMessageReceived: string; DialogEventReceived: string }
  AsyncResultStatus: { Succeeded: string }
}
export interface OfficeTeamLoginDialogOptions {
  addinOrigin: string
  authorizationEndpoint: string
  redirectUri: string
  clientId: string
  runtime?: OfficeTeamLoginDialogRuntime
  timeoutMs?: number
}
function invalid(): never {
  throw Error('team_auth_dialog_invalid')
}
const object = (value: unknown): Record<string, unknown> | undefined =>
  value !== null && typeof value === 'object' && !Array.isArray(value)
    ? (value as Record<string, unknown>)
    : undefined
const bounded = (value: unknown): value is string =>
  typeof value === 'string' &&
  new TextEncoder().encode(value).length <= 4096 &&
  !Array.from(value).some(
    (character) => character.charCodeAt(0) <= 32 || character.charCodeAt(0) === 127,
  )
function https(value: string): URL {
  let url: URL
  try {
    url = new URL(value)
  } catch {
    return invalid()
  }
  if (url.protocol !== 'https:' || url.username || url.password) return invalid()
  return url
}
function parameters(url: URL, allowed: readonly string[]): void {
  const names = Array.from(url.searchParams.keys())
  if (names.some((name) => !allowed.includes(name)) || new Set(names).size !== names.length)
    invalid()
}
export function validateTeamAuthorizationUrl(
  value: string,
  config: Pick<OfficeTeamLoginDialogOptions, 'authorizationEndpoint' | 'redirectUri' | 'clientId'>,
): string {
  if (!bounded(value)) return invalid()
  const url = https(value),
    endpoint = https(config.authorizationEndpoint)
  if (
    endpoint.search ||
    endpoint.hash ||
    url.origin !== endpoint.origin ||
    url.pathname !== endpoint.pathname ||
    url.hash
  )
    return invalid()
  parameters(url, ['client_id', 'redirect_uri', 'response_type', 'state', 'scope'])
  if (
    url.searchParams.get('client_id') !== config.clientId ||
    !config.clientId ||
    url.searchParams.get('redirect_uri') !== config.redirectUri ||
    url.searchParams.get('response_type') !== 'code' ||
    !/^[A-Za-z0-9_-]{43,128}$/.test(url.searchParams.get('state') ?? '') ||
    !url.searchParams.get('scope')
  )
    return invalid()
  return value
}
export function validateTeamCallbackUrl(value: string, redirectUri: string): string {
  if (!bounded(value)) return invalid()
  const url = https(value),
    redirect = https(redirectUri)
  if (
    redirect.search ||
    redirect.hash ||
    url.origin !== redirect.origin ||
    url.pathname !== redirect.pathname ||
    url.hash
  )
    return invalid()
  parameters(url, ['code', 'state', 'iss'])
  if (
    !url.searchParams.get('code') ||
    !/^[A-Za-z0-9_-]{43,128}$/.test(url.searchParams.get('state') ?? '')
  )
    return invalid()
  return value
}
interface Dialog {
  close(): void
  addEventHandler(name: string, callback: (event: unknown) => void): void
}
function dialog(value: unknown): Dialog | undefined {
  const candidate = object(value)
  return candidate &&
    typeof candidate.close === 'function' &&
    typeof candidate.addEventHandler === 'function'
    ? (value as Dialog)
    : undefined
}
/** Returns a verified callback URL; only the caller's authorization transaction may exchange it. */
export function createOfficeTeamLoginDialog(options: OfficeTeamLoginDialogOptions) {
  let busy = false
  return async (authorizationUrl: string, signal?: AbortSignal): Promise<string> => {
    if (busy) throw Error('team_auth_dialog_busy')
    if (signal?.aborted) throw Error('team_auth_dialog_cancelled')
    const origin = https(options.addinOrigin),
      redirect = https(options.redirectUri)
    if (
      origin.origin !== options.addinOrigin ||
      redirect.origin !== origin.origin ||
      redirect.search ||
      redirect.hash
    )
      invalid()
    validateTeamAuthorizationUrl(authorizationUrl, options)
    const runtime =
      options.runtime ??
      (typeof Office !== 'undefined'
        ? (Office as unknown as OfficeTeamLoginDialogRuntime)
        : undefined)
    try {
      if (
        !runtime?.context.requirements.isSetSupported('DialogOrigin', '1.1') ||
        typeof runtime.context.ui.displayDialogAsync !== 'function'
      )
        throw Error('team_auth_dialog_unavailable')
    } catch {
      throw Error('team_auth_dialog_unavailable')
    }
    const timeoutMs = options.timeoutMs ?? 120_000
    if (!Number.isSafeInteger(timeoutMs) || timeoutMs < 1 || timeoutMs > 600_000) invalid()
    const start = new URL('/team-auth-start.html', origin)
    start.hash = encodeURIComponent(authorizationUrl)
    busy = true
    return new Promise<string>((resolve, reject) => {
      let settled = false,
        acquired: Dialog | undefined
      const closed = new WeakSet<object>()
      const close = (value: Dialog): boolean => {
        if (closed.has(value)) return true
        closed.add(value)
        try {
          value.close()
          return true
        } catch {
          return false
        }
      }
      const finish = (code?: string, value?: string) => {
        if (settled) return
        settled = true
        clearTimeout(timer)
        signal?.removeEventListener('abort', abort)
        const didClose = !acquired || close(acquired)
        busy = false
        if (code) reject(Error(code))
        else if (!didClose) reject(Error('team_auth_dialog_failed'))
        else resolve(value!)
      }
      const abort = () => finish('team_auth_dialog_cancelled')
      const timer = setTimeout(() => finish('team_auth_dialog_timeout'), timeoutMs)
      signal?.addEventListener('abort', abort, { once: true })
      try {
        runtime.context.ui.displayDialogAsync(
          start.href,
          { displayInIframe: false, height: 60, width: 40 },
          (result) => {
            const response = object(result),
              opened = dialog(response?.value)
            if (settled) {
              if (opened) close(opened)
              return
            }
            if (acquired) {
              if (opened && opened !== acquired) close(opened)
              return
            }
            if (response?.status !== runtime.AsyncResultStatus.Succeeded || !opened) {
              if (opened) close(opened)
              finish('team_auth_dialog_failed')
              return
            }
            acquired = opened
            try {
              opened.addEventHandler(runtime.EventType.DialogMessageReceived, (event) => {
                if (settled) return
                const received = object(event)
                const messageText = received?.message
                if (
                  received?.origin !== origin.origin ||
                  typeof messageText !== 'string' ||
                  new TextEncoder().encode(messageText).length > 8192
                ) {
                  finish('team_auth_dialog_invalid')
                  return
                }
                try {
                  const message = object(JSON.parse(messageText))
                  if (
                    !message ||
                    Object.keys(message).sort().join(',') !== 'type,url' ||
                    message.type !== 'team.auth.callback' ||
                    typeof message.url !== 'string'
                  )
                    invalid()
                  const callback = validateTeamCallbackUrl(message.url, options.redirectUri)
                  if (
                    new URL(callback).searchParams.get('state') !==
                    new URL(authorizationUrl).searchParams.get('state')
                  )
                    invalid()
                  finish(undefined, callback)
                } catch {
                  finish('team_auth_dialog_invalid')
                }
              })
              opened.addEventHandler(runtime.EventType.DialogEventReceived, (event) =>
                finish(
                  object(event)?.error === 12006
                    ? 'team_auth_dialog_cancelled'
                    : 'team_auth_dialog_failed',
                ),
              )
            } catch {
              finish('team_auth_dialog_failed')
            }
          },
        )
      } catch {
        finish('team_auth_dialog_failed')
      }
      if (signal?.aborted) abort()
    })
  }
}
