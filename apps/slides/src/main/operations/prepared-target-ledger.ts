import type { PresentationTargetPreparation } from '../../shared/ipc'
import type { PresentationTargetEnrollment } from './desktop-host'

export interface PreparedTargetRequestKey {
  slideIndex: number
  sourceId?: string
}

interface PreparedTargetItem {
  request: PreparedTargetRequestKey
  response: PresentationTargetPreparation
  enrollment?: PresentationTargetEnrollment
}

interface PreparedTargetEntry {
  items: PreparedTargetItem[]
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
    const item = entry.items.find(
      (candidate) =>
        candidate.request.slideIndex === request.slideIndex &&
        candidate.request.sourceId === request.sourceId,
    )
    if (!item) return entry.completed ? { status: 'conflict', code: 'target_stale' } : undefined
    entry.touchedAt = Date.now()
    return item.response
  }

  set(
    transactionId: string,
    request: PreparedTargetRequestKey,
    response: PresentationTargetPreparation,
    enrollment?: PresentationTargetEnrollment,
  ): boolean {
    const now = Date.now()
    this.prune(now)
    const existing = this.entries.get(transactionId)
    if (existing) {
      if (existing.completed || existing.items.length >= 50) return false
      if (
        existing.items.some(
          (item) =>
            item.request.slideIndex === request.slideIndex &&
            item.request.sourceId === request.sourceId,
        )
      )
        return false
      existing.items.push({ request, response, ...(enrollment ? { enrollment } : {}) })
      existing.touchedAt = now
      existing.expiresAt = now + this.ttlMs
      return true
    }
    if (this.entries.size >= this.capacity) {
      const completed = [...this.entries.entries()]
        .filter(([, entry]) => entry.completed)
        .sort((left, right) => left[1].touchedAt - right[1].touchedAt)[0]
      if (completed) this.entries.delete(completed[0])
    }
    if (this.entries.size >= this.capacity) return false
    this.entries.set(transactionId, {
      items: [{ request, response, ...(enrollment === undefined ? {} : { enrollment }) }],
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
    for (const item of entry.items) delete item.enrollment
    entry.touchedAt = Date.now()
  }

  cancel(transactionId: string): void {
    this.complete(transactionId)
  }

  enrollment(elementId: string): PresentationTargetEnrollment | undefined {
    this.prune()
    for (const entry of this.entries.values()) {
      for (const item of entry.items) {
        if (item.enrollment?.elementId === elementId) return item.enrollment
      }
    }
    return undefined
  }
}
