export type AiProviderErrorCode =
  | 'auth_required'
  | 'model_credentials_missing'
  | 'model_rate_limited'
  | 'model_upstream_unavailable'
  | 'model_invalid_response'

export class AiProviderError extends Error {
  readonly code: AiProviderErrorCode
  readonly status: number | undefined

  constructor(code: AiProviderErrorCode, status?: number) {
    super(code)
    this.name = 'AiProviderError'
    this.code = code
    this.status = status
  }
}

/** Preserve the auth package's stable code without coupling this provider package to Electron auth. */
export function isAuthRequiredError(error: unknown): boolean {
  return Boolean(
    error &&
    typeof error === 'object' &&
    'code' in error &&
    (error as { code?: unknown }).code === 'auth_required',
  )
}

export function safeHttpProviderError(status: number): AiProviderError {
  if (status === 401 || status === 403) {
    return new AiProviderError('model_credentials_missing', status)
  }
  if (status === 429) return new AiProviderError('model_rate_limited', status)
  return new AiProviderError('model_upstream_unavailable', status)
}
