import type { ConversionKind, ConversionOutput } from './conversion-engine.js'
import { InMemoryVfs } from './vfs.js'

interface WorkerResponse {
  id: string
  ok: boolean
  outputs?: ConversionOutput[]
  error?: string
}

export interface ConversionWorkerLike {
  onmessage: ((event: MessageEvent<WorkerResponse>) => void) | null
  onerror: ((event: ErrorEvent) => void) | null
  postMessage(message: unknown, transfer?: Transferable[]): void
  terminate(): void
}

export interface ConversionRuntimeOptions {
  timeoutMs?: number
  workerFactory?: () => ConversionWorkerLike
}

export class ConversionWorkerRuntime {
  readonly #timeoutMs: number
  readonly #workerFactory: () => ConversionWorkerLike

  constructor(
    private readonly vfs: InMemoryVfs,
    options: ConversionRuntimeOptions = {},
  ) {
    this.#timeoutMs = options.timeoutMs ?? 15_000
    this.#workerFactory =
      options.workerFactory ??
      (() => new Worker(new URL('./conversion-worker.js', import.meta.url), { type: 'module' }))
  }

  run(kind: ConversionKind, inputPath: string, signal?: AbortSignal): Promise<string[]> {
    if (signal?.aborted) return Promise.reject(new Error('cancelled'))
    const normalized = this.vfs.normalize(inputPath)
    if (!normalized.startsWith('/home/user/')) return Promise.reject(new Error('vfs_path_denied'))
    const bytes = this.vfs.readBytes(normalized)
    const worker = this.#workerFactory()
    const id = crypto.randomUUID()
    return new Promise((resolve, reject) => {
      let settled = false
      const finish = (error?: string, outputs?: ConversionOutput[]) => {
        if (settled) return
        settled = true
        clearTimeout(timer)
        signal?.removeEventListener('abort', cancel)
        worker.terminate()
        if (error) return reject(new Error(error))
        try {
          if (!outputs?.length) throw new Error('conversion_invalid_output')
          const entries = outputs.map((output) => {
            if (
              !/^[A-Za-z0-9._-]{1,128}$/.test(output.path) ||
              !(output.bytes instanceof Uint8Array)
            )
              throw new Error('conversion_invalid_output')
            return [`/home/user/${output.path}`, output.bytes] as const
          })
          this.vfs.writeBatch(entries)
          resolve(entries.map(([path]) => path))
        } catch (caught) {
          reject(caught instanceof Error ? caught : new Error('conversion_invalid_output'))
        }
      }
      const cancel = () => finish('cancelled')
      const timer = setTimeout(() => finish('command_timeout'), this.#timeoutMs)
      signal?.addEventListener('abort', cancel, { once: true })
      worker.onerror = () => finish('conversion_failed')
      worker.onmessage = (event) => {
        if (event.data?.id !== id) return
        finish(
          event.data.ok ? undefined : event.data.error || 'conversion_failed',
          event.data.outputs,
        )
      }
      const copy = bytes.slice()
      worker.postMessage({ id, kind, inputName: normalized.split('/').pop(), bytes: copy }, [
        copy.buffer,
      ])
    })
  }
}
