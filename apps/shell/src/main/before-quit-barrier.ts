export interface BeforeQuitEventLike {
  preventDefault(): void
}

export function createBeforeQuitBarrier(options: {
  readonly cleanup: () => Promise<void>
  readonly quit: () => void
  readonly deadlineMs?: number
  readonly diagnostics?: (code: string) => void
}): (event: BeforeQuitEventLike) => void {
  const deadlineMs = options.deadlineMs ?? 5_000
  if (!Number.isSafeInteger(deadlineMs) || deadlineMs <= 0)
    throw new TypeError('invalid_quit_deadline')
  let started = false
  let complete = false
  let quitSent = false
  return (event) => {
    if (complete) return
    event.preventDefault()
    if (started) return
    started = true
    const deadlineTimer = setTimeout(
      () => options.diagnostics?.('enhanced_quit_cleanup_deadline'),
      deadlineMs,
    )
    deadlineTimer.unref()
    let cleanup: Promise<void>
    try {
      cleanup = options.cleanup()
    } catch {
      clearTimeout(deadlineTimer)
      options.diagnostics?.('enhanced_quit_cleanup_failed')
      return
    }
    void cleanup
      .then(() => {
        clearTimeout(deadlineTimer)
        complete = true
        if (!quitSent) {
          quitSent = true
          options.quit()
        }
      })
      .catch(() => {
        clearTimeout(deadlineTimer)
        options.diagnostics?.('enhanced_quit_cleanup_failed')
      })
  }
}
