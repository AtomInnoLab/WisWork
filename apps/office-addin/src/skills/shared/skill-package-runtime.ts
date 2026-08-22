import type { ParsedSkillArchive } from './skill-package.js'

interface PackageWorkerResponse {
  id: string
  ok: boolean
  pkg?: ParsedSkillArchive
  error?: string
}

export interface PackageWorkerLike {
  onmessage: ((event: MessageEvent<PackageWorkerResponse>) => void) | null
  onerror: ((event: ErrorEvent) => void) | null
  postMessage(message: unknown, transfer?: Transferable[]): void
  terminate(): void
}

export class SkillPackageWorkerRuntime {
  readonly #workerFactory: () => PackageWorkerLike
  readonly #timeoutMs: number
  readonly #active = new Set<(code: string) => void>()

  constructor(options: { workerFactory?: () => PackageWorkerLike; timeoutMs?: number } = {}) {
    this.#workerFactory =
      options.workerFactory ??
      (() => new Worker(new URL('./skill-package-worker.js', import.meta.url), { type: 'module' }))
    this.#timeoutMs = options.timeoutMs ?? 10_000
  }

  parse(source: ArrayBuffer, signal?: AbortSignal): Promise<ParsedSkillArchive> {
    if (signal?.aborted) return Promise.reject(new Error('upload_cancelled'))
    const worker = this.#workerFactory()
    const id = crypto.randomUUID()
    return new Promise((resolve, reject) => {
      let settled = false
      const finish = (error?: string, pkg?: ParsedSkillArchive) => {
        if (settled) return
        settled = true
        clearTimeout(timer)
        signal?.removeEventListener('abort', cancel)
        this.#active.delete(finish)
        worker.terminate()
        if (error) reject(new Error(error))
        else if (pkg) resolve(pkg)
        else reject(new Error('invalid_skill_package'))
      }
      const cancel = () => finish('upload_cancelled')
      const timer = setTimeout(() => finish('skill_package_timeout'), this.#timeoutMs)
      this.#active.add(finish)
      signal?.addEventListener('abort', cancel, { once: true })
      worker.onerror = () => finish('invalid_skill_package')
      worker.onmessage = (event) => {
        if (event.data?.id !== id) return
        finish(
          event.data.ok ? undefined : event.data.error || 'invalid_skill_package',
          event.data.pkg,
        )
      }
      const bytes = new Uint8Array(source.slice(0))
      let posted = false
      try {
        worker.postMessage({ id, bytes }, [bytes.buffer])
        posted = true
      } catch {
        // The stable failure is finalized below so termination cannot be skipped.
      } finally {
        if (!posted) finish('invalid_skill_package')
      }
    })
  }

  cancelAll(): void {
    for (const finish of [...this.#active]) finish('upload_cancelled')
  }
}
