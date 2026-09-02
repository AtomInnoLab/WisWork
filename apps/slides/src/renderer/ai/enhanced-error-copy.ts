export function friendlyEnhancedError(
  error: string,
  genericFailure: string,
  timeoutFailure: string,
): string {
  if (error === 'enhanced_turn_timeout') return timeoutFailure
  if (/^enhanced_[a-z0-9_]+$/.test(error)) return genericFailure
  return error
}

export function shouldMarkEnhancedMessageUndelivered(executedToolCount: number): boolean {
  return executedToolCount === 0
}
