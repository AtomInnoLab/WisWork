export const OFFICE_WRITE_READBACK_ATTEMPTS = 3
export const OFFICE_WRITE_READBACK_DELAY_MS = 50

function cancelled(signal?: AbortSignal): void {
  if (signal?.aborted) throw new Error('cancelled')
}

async function wait(milliseconds: number, signal?: AbortSignal): Promise<void> {
  cancelled(signal)
  await new Promise<void>((resolve, reject) => {
    const timer = setTimeout(resolve, milliseconds)
    signal?.addEventListener(
      'abort',
      () => {
        clearTimeout(timer)
        reject(new Error('cancelled'))
      },
      { once: true },
    )
  })
  cancelled(signal)
}

export async function readUntilConverged<T>(options: {
  read(): Promise<T>
  accept(value: T): boolean
  signal?: AbortSignal
  attempts?: number
  delayMs?: number
  delay?: (milliseconds: number, signal?: AbortSignal) => Promise<void>
}): Promise<T> {
  const attempts = options.attempts ?? OFFICE_WRITE_READBACK_ATTEMPTS
  if (!Number.isSafeInteger(attempts) || attempts < 1) throw new Error('invalid_tool_input')
  let value!: T
  for (let attempt = 0; attempt < attempts; attempt += 1) {
    cancelled(options.signal)
    value = await options.read()
    cancelled(options.signal)
    if (options.accept(value) || attempt === attempts - 1) return value
    await (options.delay ?? wait)(options.delayMs ?? OFFICE_WRITE_READBACK_DELAY_MS, options.signal)
  }
  return value
}
