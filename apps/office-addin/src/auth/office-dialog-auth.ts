export interface OfficeAuthDialog {
  close(): void
  onMessage(listener: (message: string) => void): () => void
  onError(listener: () => void): () => void
}

export interface OfficeDialogRuntime {
  open(url: string): Promise<OfficeAuthDialog>
}

interface DialogAuthClient {
  startAuthorization(): Promise<string>
  consumeCallback(url: string): Promise<void>
  cancelAuthorization(): void
}

const CALLBACK_MESSAGE_TYPE = 'wiswork_oauth_callback'
const MAX_CALLBACK_URL_LENGTH = 4_096
const DIALOG_START_PATH = '/oauth/dialog-start.html'

export function createDialogStartUrl(authorizationUrl: string, callbackUrl: string): string {
  const start = new URL(DIALOG_START_PATH, new URL(callbackUrl).origin)
  start.hash = encodeURIComponent(authorizationUrl)
  return start.toString()
}

export function authorizationFromDialogStart(
  currentUrl: string,
  expectedAuthorizationUrl: string,
): string | undefined {
  try {
    const current = new URL(currentUrl)
    if (current.pathname !== DIALOG_START_PATH || current.search || !current.hash) return undefined
    const authorization = new URL(decodeURIComponent(current.hash.slice(1)))
    const expected = new URL(expectedAuthorizationUrl)
    if (
      authorization.origin !== expected.origin ||
      authorization.pathname !== expected.pathname ||
      authorization.username ||
      authorization.password ||
      authorization.hash
    )
      return undefined
    return authorization.toString()
  } catch {
    return undefined
  }
}

function exactCallbackUrl(rawUrl: string, expectedUrl: string): boolean {
  if (rawUrl.length > MAX_CALLBACK_URL_LENGTH) return false
  try {
    const actual = new URL(rawUrl)
    const expected = new URL(expectedUrl)
    return (
      actual.origin === expected.origin &&
      actual.pathname === expected.pathname &&
      !actual.username &&
      !actual.password &&
      !actual.hash
    )
  } catch {
    return false
  }
}

function callbackFromMessage(message: string, expectedUrl: string): string | undefined {
  if (message.length > MAX_CALLBACK_URL_LENGTH + 128) return undefined
  try {
    const value: unknown = JSON.parse(message)
    if (!value || typeof value !== 'object' || Array.isArray(value)) return undefined
    const record = value as Record<string, unknown>
    if (
      Object.keys(record).length !== 2 ||
      record.type !== CALLBACK_MESSAGE_TYPE ||
      typeof record.url !== 'string' ||
      !exactCallbackUrl(record.url, expectedUrl)
    )
      return undefined
    return record.url
  } catch {
    return undefined
  }
}

export function createOfficeDialogAuth(runtime: OfficeDialogRuntime, callbackUrl: string) {
  return async (auth: DialogAuthClient): Promise<void> => {
    const authorizationUrl = await auth.startAuthorization()
    let dialog: OfficeAuthDialog
    try {
      dialog = await runtime.open(authorizationUrl)
    } catch (error) {
      auth.cancelAuthorization()
      throw error
    }
    await new Promise<void>((resolve, reject) => {
      let settled = false
      let removeMessage: () => void = () => undefined
      let removeError: () => void = () => undefined
      const finish = (result: () => void) => {
        if (settled) return
        settled = true
        removeMessage()
        removeError()
        dialog.close()
        auth.cancelAuthorization()
        result()
      }
      removeMessage = dialog.onMessage((message) => {
        const callback = callbackFromMessage(message, callbackUrl)
        if (!callback) {
          finish(() => reject(new Error('invalid_callback')))
          return
        }
        if (settled) return
        settled = true
        removeMessage()
        removeError()
        dialog.close()
        void auth.consumeCallback(callback).then(resolve, reject)
      })
      removeError = dialog.onError(() => finish(() => reject(new Error('sign_in_failed'))))
    })
  }
}

export function forwardOAuthCallbackToParent(
  currentUrl: string,
  callbackUrl: string,
  messageParent: (message: string, targetOrigin: string) => void,
): boolean {
  if (!exactCallbackUrl(currentUrl, callbackUrl)) return false
  try {
    messageParent(
      JSON.stringify({ type: CALLBACK_MESSAGE_TYPE, url: currentUrl }),
      new URL(callbackUrl).origin,
    )
    return true
  } catch {
    return false
  }
}

interface BrowserOfficeDialogDependencies {
  supportsDialogOrigin?: () => boolean
}

export function dialogMessageOriginIsTrusted(
  origin: string | undefined,
  trustedOrigin: string,
  supportsDialogOrigin: boolean,
): boolean {
  return origin === trustedOrigin || (!supportsDialogOrigin && origin === undefined)
}

export function createBrowserOfficeDialogRuntime(
  callbackUrl: string,
  dependencies: BrowserOfficeDialogDependencies = {},
): OfficeDialogRuntime {
  const trustedCallbackOrigin = new URL(callbackUrl).origin
  const supportsDialogOrigin =
    dependencies.supportsDialogOrigin ??
    (() => Office.context.requirements.isSetSupported('DialogOrigin', '1.1'))
  return {
    open(url) {
      const hasDialogOrigin = supportsDialogOrigin()
      return new Promise((resolve, reject) => {
        Office.context.ui.displayDialogAsync(
          createDialogStartUrl(url, callbackUrl),
          { height: 60, width: 40, displayInIframe: false },
          (result) => {
            if (result.status !== Office.AsyncResultStatus.Succeeded) {
              reject(new Error('sign_in_failed'))
              return
            }
            const dialog = result.value
            resolve({
              close: () => dialog.close(),
              onMessage(listener) {
                const handler = (
                  event: { message: string; origin: string | undefined } | { error: number },
                ) => {
                  if (
                    'message' in event &&
                    dialogMessageOriginIsTrusted(
                      event.origin,
                      trustedCallbackOrigin,
                      hasDialogOrigin,
                    )
                  )
                    listener(event.message)
                }
                dialog.addEventHandler(Office.EventType.DialogMessageReceived, handler)
                return () => undefined
              },
              onError(listener) {
                const handler = () => listener()
                dialog.addEventHandler(Office.EventType.DialogEventReceived, handler)
                return () => undefined
              },
            })
          },
        )
      })
    },
  }
}

export function forwardBrowserOAuthCallback(currentUrl: string, callbackUrl: string): boolean {
  const hasDialogOrigin = Office.context.requirements.isSetSupported('DialogOrigin', '1.1')
  return forwardOAuthCallbackToParent(currentUrl, callbackUrl, (message, targetOrigin) => {
    if (hasDialogOrigin) Office.context.ui.messageParent(message, { targetOrigin })
    else Office.context.ui.messageParent(message)
  })
}
