import type { AccountLoginEvent } from '../../shared/home-api'

const SAFE_ERROR_CODES = new Set([
  'invalid_response',
  'invalid_callback',
  'invalid_state',
  'callback_reused',
  'callback_expired',
  'secure_storage_unavailable',
  'auth_required',
  'network_error',
  'auth_not_initialized',
  'login_failed',
])

export function loginErrorKind(error: string | undefined): 'network' | 'expired' | 'failed' {
  if (error === 'network_error') return 'network'
  if (error === 'callback_expired') return 'expired'
  return 'failed'
}

export function formatAccountLoginDiagnostic(event: AccountLoginEvent): string {
  const error = event.error
  if (!error || !SAFE_ERROR_CODES.has(error)) return 'login_failed'
  const parts = [error]
  const diagnostic = event.diagnostic
  if (diagnostic && (diagnostic.stage === 'callback_exchange' || diagnostic.stage === 'refresh')) {
    parts.push(diagnostic.stage)
    if (
      diagnostic.httpStatus !== undefined &&
      Number.isInteger(diagnostic.httpStatus) &&
      diagnostic.httpStatus >= 100 &&
      diagnostic.httpStatus <= 599
    )
      parts.push(`HTTP ${diagnostic.httpStatus}`)
  }
  return parts.join(' · ')
}
