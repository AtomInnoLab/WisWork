import { describe, expect, it, vi } from 'vitest'
import { captureAndScrubOAuthCallback } from '../src/auth/oauth-callback.js'

const callbackUrl = 'https://localhost:3000/oauth/callback'

describe('OAuth callback location handling', () => {
  it.each([
    '?code=secret-code&state=wrong',
    '?code=one&code=two&state=s',
    '?error=access_denied&state=s',
  ])('captures then immediately scrubs callback parameters for %s', (query) => {
    const replace = vi.fn()
    const href = `${callbackUrl}${query}`

    expect(captureAndScrubOAuthCallback(callbackUrl, href, replace)).toBe(href)
    expect(replace).toHaveBeenCalledOnce()
    expect(replace).toHaveBeenCalledWith('https://localhost:3000/taskpane.html')
  })

  it('does not alter a non-callback page', () => {
    const replace = vi.fn()
    expect(
      captureAndScrubOAuthCallback(callbackUrl, 'https://localhost:3000/taskpane.html', replace),
    ).toBeUndefined()
    expect(replace).not.toHaveBeenCalled()
  })
})
