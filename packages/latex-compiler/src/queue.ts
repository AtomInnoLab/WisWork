import { LatexCompilerError } from './errors.js'

export interface CompileQueueRequest<T> {
  readonly projectId: string
  readonly revision: string
  readonly run: (context: { signal: AbortSignal }) => Promise<T>
  readonly publish?: (value: T) => void | Promise<void>
}

interface ActiveCompile<T> {
  readonly revision: string
  readonly controller: AbortController
  readonly promise: Promise<T>
  readonly token: symbol
  phase: 'running' | 'publishing'
  cancelled: boolean
}

export class CompileQueue<T> {
  private readonly active = new Map<string, ActiveCompile<T>>()

  request(request: CompileQueueRequest<T>): Promise<T> {
    const existing = this.active.get(request.projectId)
    if (existing?.revision === request.revision) return existing.promise
    if (existing) {
      if (existing.phase === 'running') {
        existing.cancelled = true
        existing.controller.abort()
      }
      return existing.promise.catch(() => undefined).then(() => this.request(request))
    }

    const controller = new AbortController()
    const token = Symbol(request.revision)
    const isCurrent = () => this.active.get(request.projectId)?.token === token
    const stale = () => new LatexCompilerError('TECTONIC_STALE_RESULT', 'Stale compile result')
    const promise = Promise.resolve()
      .then(() => request.run({ signal: controller.signal }))
      .then(async (value) => {
        const active = this.active.get(request.projectId)
        if (!active || active.token !== token || active.cancelled) throw stale()
        active.phase = 'publishing'
        await request.publish?.(value)
        if (!isCurrent()) throw stale()
        return value
      })
      .finally(() => {
        if (isCurrent()) this.active.delete(request.projectId)
      })
    this.active.set(request.projectId, {
      revision: request.revision,
      controller,
      promise,
      token,
      phase: 'running',
      cancelled: false,
    })
    return promise
  }

  cancel(projectId: string): boolean {
    const active = this.active.get(projectId)
    if (!active || active.phase === 'publishing' || active.cancelled) return false
    active.cancelled = true
    active.controller.abort()
    return true
  }
}
