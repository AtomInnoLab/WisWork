import { AuthError, extractCallbackUrl } from '@wiswork/auth'
import type { AccountLoginEvent } from '../shared/home-api'

export interface AuthDeepLinkQueueOptions {
  notify(event: AccountLoginEvent): void
}

export function createAuthDeepLinkQueue(options: AuthDeepLinkQueueOptions) {
  let consume: ((url: string) => Promise<unknown>) | null = null
  const pending: string[] = []
  let tail: Promise<void> = Promise.resolve()

  const enqueue = (url: string) => {
    const activeConsumer = consume
    if (!activeConsumer) {
      pending.push(url)
      return
    }
    tail = tail.then(async () => {
      try {
        await activeConsumer(url)
        options.notify({ phase: 'success' })
      } catch (error) {
        const diagnostic =
          error instanceof AuthError &&
          error.diagnostic &&
          (error.diagnostic.stage === 'callback_exchange' ||
            error.diagnostic.stage === 'refresh') &&
          (error.diagnostic.httpStatus === undefined ||
            (Number.isInteger(error.diagnostic.httpStatus) &&
              error.diagnostic.httpStatus >= 100 &&
              error.diagnostic.httpStatus <= 599))
            ? error.diagnostic
            : undefined
        const event: AccountLoginEvent = {
          phase: 'error',
          error: error instanceof AuthError ? error.code : 'login_failed',
          ...(diagnostic ? { diagnostic } : {}),
        }
        console.warn('[auth] login callback failed', event)
        options.notify(event)
      }
    })
  }

  return {
    handle(input: string | readonly string[]): boolean {
      const callback = extractCallbackUrl(input)
      if (!callback) return false
      enqueue(callback)
      return true
    },
    async initialize(next: (url: string) => Promise<unknown>): Promise<void> {
      if (!consume) {
        consume = next
        for (const callback of pending.splice(0)) enqueue(callback)
      }
      await tail
    },
    whenIdle(): Promise<void> {
      return tail
    },
  }
}
