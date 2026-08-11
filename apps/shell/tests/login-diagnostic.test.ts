import { describe, expect, it } from 'vitest'
import { formatAccountLoginDiagnostic, loginErrorKind } from '../src/renderer/src/login-diagnostic'

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
})
