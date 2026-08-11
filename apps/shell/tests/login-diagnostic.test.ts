import { describe, expect, it } from 'vitest'
import { createElement } from 'react'
import { renderToStaticMarkup } from 'react-dom/server'
import { formatAccountLoginDiagnostic, loginErrorKind } from '../src/renderer/src/login-diagnostic'
import { LoginDiagnostic } from '../src/renderer/src/LoginDiagnostic'

describe('login diagnostics', () => {
  it('maps stable auth codes and displays only bounded safe metadata', () => {
    const event = {
      phase: 'error' as const,
      error: 'network_error',
      diagnostic: { stage: 'callback_exchange' as const, httpStatus: 400 },
    }
    expect(loginErrorKind(event.error)).toBe('network')
    expect(formatAccountLoginDiagnostic(event)).toBe('network_error · callback_exchange · HTTP 400')
    expect(loginErrorKind('callback_expired')).toBe('expired')
    expect(
      formatAccountLoginDiagnostic({
        phase: 'error',
        error: 'private token value',
        diagnostic: { stage: 'callback_exchange', httpStatus: 99999 },
      }),
    ).toBe('login_failed')
  })

  it('renders the complete diagnostic in a copyable account-menu panel', () => {
    const value = 'network_error · callback_exchange · HTTP 400'
    const html = renderToStaticMarkup(
      createElement(LoginDiagnostic, { value, onCopy: () => undefined }),
    )
    expect(html).toContain(value)
    expect(html).toContain('复制诊断 / Copy')
    expect(html).toContain('login-diagnostic-value')
  })
})
