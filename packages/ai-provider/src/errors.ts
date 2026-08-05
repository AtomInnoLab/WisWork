export type AiProviderErrorCode =
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

export function safeHttpProviderError(status: number): AiProviderError {
  if (status === 401 || status === 403) {
    return new AiProviderError('model_credentials_missing', status)
  }
  if (status === 429) return new AiProviderError('model_rate_limited', status)
  return new AiProviderError('model_upstream_unavailable', status)
}
