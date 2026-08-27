import type { PresentationTargetPreparation } from '../../shared/ipc'
import type { PresentationTargetEnrollment } from './desktop-host'

export interface PreparedTargetRequestKey {
  slideIndex: number
  sourceId?: string
}

interface PreparedTargetEntry {
  request: PreparedTargetRequestKey
  response: PresentationTargetPreparation
  enrollment?: PresentationTargetEnrollment
  expiresAt: number
  completed: boolean
  touchedAt: number
}

export class PreparedTargetLedger {
  private readonly entries = new Map<string, PreparedTargetEntry>()

  constructor(
    private readonly capacity = 64,
    private readonly ttlMs = 5 * 60_000,
  ) {}

  private prune(now = Date.now()): void {
    for (const [id, entry] of this.entries) {
      if (entry.expiresAt <= now) this.entries.delete(id)
    }
  }

  get(
    transactionId: string,
    request: PreparedTargetRequestKey,
  ): PresentationTargetPreparation | undefined {
    this.prune()
    const entry = this.entries.get(transactionId)
    if (!entry) return undefined
    if (
      entry.request.slideIndex !== request.slideIndex ||
      entry.request.sourceId !== request.sourceId
    )
      return { status: 'conflict', code: 'target_stale' }
    entry.touchedAt = Date.now()
    return entry.response
  }

  set(
    transactionId: string,
    request: PreparedTargetRequestKey,
    response: PresentationTargetPreparation,
    enrollment?: PresentationTargetEnrollment,
  ): boolean {
    const now = Date.now()
    this.prune(now)
    if (this.entries.has(transactionId)) return false
    if (this.entries.size >= this.capacity) {
      const completed = [...this.entries.entries()]
        .filter(([, entry]) => entry.completed)
        .sort((left, right) => left[1].touchedAt - right[1].touchedAt)[0]
      if (completed) this.entries.delete(completed[0])
    }
    if (this.entries.size >= this.capacity) return false
    this.entries.set(transactionId, {
      request,
      response,
      ...(enrollment === undefined ? {} : { enrollment }),
      expiresAt: now + this.ttlMs,
      completed: false,
      touchedAt: now,
    })
    return true
  }

  complete(transactionId: string): void {
    const entry = this.entries.get(transactionId)
    if (!entry) return
    entry.completed = true
    delete entry.enrollment
    entry.touchedAt = Date.now()
  }

  cancel(transactionId: string): void {
    this.complete(transactionId)
  }

  enrollment(elementId: string): PresentationTargetEnrollment | undefined {
    this.prune()
    for (const entry of this.entries.values()) {
      if (entry.enrollment?.elementId === elementId) return entry.enrollment
    }
    return undefined
  }
}
