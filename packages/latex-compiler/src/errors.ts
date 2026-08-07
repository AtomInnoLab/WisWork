export type LatexCompilerErrorCode =
  | 'TECTONIC_MANIFEST_INVALID'
  | 'BUNDLE_DOWNLOAD_CANCELLED'
  | 'BUNDLE_DOWNLOAD_FAILED'
  | 'BUNDLE_INTEGRITY_FAILED'
  | 'BUNDLE_INSTALL_FAILED'
  | 'TECTONIC_WORKSPACE_INVALID'
  | 'TECTONIC_EXIT_NONZERO'
  | 'TECTONIC_TOTAL_TIMEOUT'
  | 'TECTONIC_IDLE_TIMEOUT'
  | 'TECTONIC_OUTPUT_LIMIT'
  | 'TECTONIC_CANCELLED'
  | 'TECTONIC_STALE_RESULT'

export class LatexCompilerError extends Error {
  readonly code: LatexCompilerErrorCode
  override readonly cause?: unknown

  constructor(code: LatexCompilerErrorCode, message: string, cause?: unknown) {
    super(message)
    this.name = 'LatexCompilerError'
    this.code = code
    this.cause = cause
  }
}

export function manifestError(message: string): never {
  throw new LatexCompilerError('TECTONIC_MANIFEST_INVALID', message)
}
