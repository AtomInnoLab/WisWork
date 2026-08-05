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
        options.notify({
          phase: 'error',
          error: error instanceof AuthError ? error.code : 'login_failed',
        })
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
