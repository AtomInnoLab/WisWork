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
    const deadline = new Promise<'deadline'>((resolve) => {
      const timer = setTimeout(() => resolve('deadline'), deadlineMs)
      timer.unref()
    })
    void Promise.race([options.cleanup().then(() => 'clean' as const), deadline])
      .then((result) => {
        if (result === 'deadline') options.diagnostics?.('enhanced_quit_cleanup_deadline')
      })
      .catch(() => options.diagnostics?.('enhanced_quit_cleanup_failed'))
      .finally(() => {
        complete = true
        if (!quitSent) {
          quitSent = true
          options.quit()
        }
      })
  }
}
