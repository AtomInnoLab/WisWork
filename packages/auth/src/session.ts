export interface AuthSession {
  accessToken: string
  refreshToken: string
  userId: string
  email?: string
  expiresAt?: number
}

export interface AccountStatus {
  loggedIn: boolean
  email?: string
  userId?: string
}

export interface SessionStore {
  load(): Promise<AuthSession | null>
  save(session: AuthSession): Promise<void>
  clear(): Promise<void>
}

export type AuthErrorCode =
  | 'invalid_response'
  | 'invalid_callback'
  | 'invalid_state'
  | 'callback_reused'
  | 'callback_expired'
  | 'secure_storage_unavailable'
  | 'auth_required'
  | 'network_error'
  | 'auth_not_initialized'
  | 'auth_unavailable_in_standalone'

export interface AuthDiagnostic {
  stage: 'callback_exchange' | 'refresh'
  httpStatus?: number
}

export class AuthError extends Error {
  constructor(
    public readonly code: AuthErrorCode,
    public readonly diagnostic?: AuthDiagnostic,
  ) {
    super(code)
    this.name = 'AuthError'
  }
}

export function publicAccountStatus(session: AuthSession | null): AccountStatus {
  if (!session) return { loggedIn: false }
  return {
    loggedIn: true,
    ...(session.email ? { email: session.email } : {}),
    ...(session.userId ? { userId: session.userId } : {}),
  }
}

export function parseSessionPayload(value: unknown, now: number): AuthSession {
  if (!value || typeof value !== 'object' || Array.isArray(value))
    throw new AuthError('invalid_response')
  const record = value as Record<string, unknown>
  const accessToken = record.token
  const refreshToken = record.refresh_token
  const rawUserId = record.user_id
  if (typeof accessToken !== 'string' || !accessToken || accessToken.length > 16_384)
    throw new AuthError('invalid_response')
  if (typeof refreshToken !== 'string' || !refreshToken || refreshToken.length > 16_384)
    throw new AuthError('invalid_response')
  const userId =
    typeof rawUserId === 'string'
      ? rawUserId
      : typeof rawUserId === 'number' && Number.isSafeInteger(rawUserId) && rawUserId >= 0
        ? String(rawUserId)
        : undefined
  if (!userId || userId.length > 512) throw new AuthError('invalid_response')
  const email =
    typeof record.email === 'string' && record.email.length <= 512 ? record.email : undefined
  const expiresIn =
    typeof record.expires_in === 'number' &&
    Number.isFinite(record.expires_in) &&
    record.expires_in > 0
      ? record.expires_in
      : undefined
  return {
    accessToken,
    refreshToken,
    userId,
    ...(email ? { email } : {}),
    ...(expiresIn ? { expiresAt: now + expiresIn * 1_000 } : {}),
  }
}
