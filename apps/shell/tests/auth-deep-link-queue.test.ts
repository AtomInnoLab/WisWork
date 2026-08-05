import { describe, expect, it, vi } from 'vitest'
import { AuthError } from '@wiswork/auth'
import { createAuthDeepLinkQueue } from '../src/main/auth-deep-link-queue'

describe('auth deep-link queue', () => {
  it('queues callbacks before initialization and consumes each exactly once in order', async () => {
    const consume = vi.fn(async () => undefined)
    const notify = vi.fn()
    const queue = createAuthDeepLinkQueue({ notify })

    expect(queue.handle('wiswork://oauth/callback?code=first&state=one')).toBe(true)
    expect(queue.handle(['wiswork', 'wiswork://oauth/callback?code=second&state=two'])).toBe(true)
    expect(consume).not.toHaveBeenCalled()

    await queue.initialize(consume)
    await queue.initialize(consume)

    expect(consume.mock.calls.map(([url]) => url)).toEqual([
      'wiswork://oauth/callback?code=first&state=one',
      'wiswork://oauth/callback?code=second&state=two',
    ])
    expect(notify).toHaveBeenCalledTimes(2)
    expect(notify).toHaveBeenCalledWith({ phase: 'success' })
  })

  it('ignores unrelated input and exposes only stable error codes', async () => {
    const notify = vi.fn()
    const queue = createAuthDeepLinkQueue({ notify })
    expect(queue.handle('https://example.com/?token=private')).toBe(false)

    await queue.initialize(async () => {
      throw new Error('private code and token')
    })
    queue.handle('wiswork://oauth/callback?code=private&state=one')
    await queue.whenIdle()
    expect(notify).toHaveBeenLastCalledWith({ phase: 'error', error: 'login_failed' })
    expect(JSON.stringify(notify.mock.calls)).not.toContain('private code and token')

    const authQueue = createAuthDeepLinkQueue({ notify })
    await authQueue.initialize(async () => {
      throw new AuthError('callback_expired')
    })
    authQueue.handle('wiswork://oauth/callback?code=private&state=two')
    await authQueue.whenIdle()
    expect(notify).toHaveBeenLastCalledWith({ phase: 'error', error: 'callback_expired' })
  })
})
