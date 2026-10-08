/** In-memory, document-bound FIFO for agent edits captured from the canvas selection. */

export const EDIT_QUEUE_MAX_TARGETS = 10
export const EDIT_QUEUE_DEFAULT_CAPACITY = 20

export interface SelectionElementScope {
  elementId: string
  expectedType?: string
  expectedFingerprint?: string
}

export interface SelectionSlideScope {
  slideId: string
  elements?: readonly SelectionElementScope[]
  /** Explicit opt-in for operations whose target is the entire selected slide. */
  allowSlideTarget?: boolean
}

export interface SelectionScope {
  documentId: string
  sessionId: string
  generation: number
  slides: readonly SelectionSlideScope[]
}

export interface SelectionEditReceipt {
  taskId: string
  invocationId: string
  transactionId?: string
  status: 'applied' | 'unchanged' | 'conflict' | 'uncertain' | 'cancelled' | 'failed'
  [key: string]: unknown
}

export interface SelectionEditTaskInput {
  invocationId: string
  instruction: string
  scope: SelectionScope
}

export interface SelectionEditRunContext {
  readonly taskId: string
  readonly invocationId: string
  readonly instruction: string
  readonly scope: SelectionScope
  readonly signal: AbortSignal
  /** Call immediately before the first host mutation is dispatched. */
  markWriteStarted(): void
}

export type SelectionEditRunner = (
  context: SelectionEditRunContext,
) => Promise<SelectionEditReceipt>

export class SelectionScopeConflict extends Error {
  readonly code = 'selection_scope_conflict'
  constructor(message = 'The edit target is outside the captured selection') {
    super(message)
    this.name = 'SelectionScopeConflict'
  }
}

export class EditQueueCapacityError extends Error {
  constructor() {
    super('Selection edit queue capacity exceeded')
    this.name = 'EditQueueCapacityError'
  }
}

type Deferred = {
  promise: Promise<SelectionEditReceipt>
  resolve: (receipt: SelectionEditReceipt) => void
  reject: (error: unknown) => void
}

interface Entry {
  taskId: string
  input: SelectionEditTaskInput
  runner: SelectionEditRunner
  deferred: Deferred
  controller: AbortController
  state: 'queued' | 'running' | 'settled'
  writeStarted: boolean
}

export interface EditQueueSnapshot {
  queued: number
  running: boolean
  paused: boolean
  activeTaskId?: string
  receipts: readonly SelectionEditReceipt[]
}

let taskSequence = 0
const nextTaskId = (): string => `selection-edit-${Date.now().toString(36)}-${++taskSequence}`

const abortError = (reason: string): DOMException => new DOMException(reason, 'AbortError')

const validateAndCopyScope = (scope: SelectionScope): SelectionScope => {
  if (!scope.documentId || !scope.sessionId || !Number.isSafeInteger(scope.generation))
    throw new SelectionScopeConflict('Selection identity is incomplete')
  const count = scope.slides.reduce(
    (sum, slide) => sum + Math.max(1, slide.elements?.length ?? 0),
    0,
  )
  if (count === 0 || count > EDIT_QUEUE_MAX_TARGETS)
    throw new SelectionScopeConflict(`Selection must contain 1-${EDIT_QUEUE_MAX_TARGETS} targets`)
  const slides = scope.slides.map((slide) => ({
    slideId: slide.slideId,
    ...(slide.elements ? { elements: slide.elements.map((element) => ({ ...element })) } : {}),
    ...(slide.allowSlideTarget ? { allowSlideTarget: true } : {}),
  }))
  if (slides.some((slide) => !slide.slideId))
    throw new SelectionScopeConflict('Slide id is missing')
  const slideIds = new Set<string>()
  const elementIds = new Set<string>()
  for (const slide of slides) {
    if (slideIds.has(slide.slideId)) throw new SelectionScopeConflict('Duplicate slide target')
    slideIds.add(slide.slideId)
    for (const element of slide.elements ?? []) {
      if (!element.elementId || elementIds.has(element.elementId))
        throw new SelectionScopeConflict('Duplicate or missing element target')
      elementIds.add(element.elementId)
    }
  }
  return Object.freeze({ ...scope, slides: Object.freeze(slides) })
}

/**
 * Structural gate used at prepare/compile time. It deliberately accepts only
 * durable target metadata and never slide XML or presentation content.
 */
export function assertOperationsWithinSelectionScope(
  scope: SelectionScope,
  operations: readonly unknown[],
): void {
  const captured = validateAndCopyScope(scope)
  for (const operation of operations) {
    if (!operation || typeof operation !== 'object' || Array.isArray(operation))
      throw new SelectionScopeConflict()
    const target = (operation as { target?: unknown }).target
    if (!target || typeof target !== 'object' || Array.isArray(target))
      throw new SelectionScopeConflict()
    const durable = target as {
      slideId?: unknown
      elementId?: unknown
      expectedType?: unknown
      expectedFingerprint?: unknown
    }
    if (typeof durable.slideId !== 'string') throw new SelectionScopeConflict()
    const slide = captured.slides.find((item) => item.slideId === durable.slideId)
    if (!slide) throw new SelectionScopeConflict()
    if (typeof durable.elementId !== 'string') {
      if (!slide.allowSlideTarget || (slide.elements?.length ?? 0) > 0)
        throw new SelectionScopeConflict()
      continue
    }
    const element = slide.elements?.find((item) => item.elementId === durable.elementId)
    if (!element) throw new SelectionScopeConflict()
    if (element.expectedType && durable.expectedType !== element.expectedType)
      throw new SelectionScopeConflict('Selected element type changed')
    if (element.expectedFingerprint && durable.expectedFingerprint !== element.expectedFingerprint)
      throw new SelectionScopeConflict('Selected element fingerprint changed')
  }
}

export function selectionScopeSummary(scope: SelectionScope): string {
  const elements = scope.slides.reduce((sum, slide) => sum + (slide.elements?.length ?? 0), 0)
  return elements > 0
    ? `${elements} selected element${elements === 1 ? '' : 's'} on ${scope.slides.length} slide${scope.slides.length === 1 ? '' : 's'}`
    : `${scope.slides.length} selected slide${scope.slides.length === 1 ? '' : 's'}`
}

export class SelectionEditQueue {
  private readonly capacity: number
  private entries: Entry[] = []
  private active: Entry | null = null
  private paused = false
  private receipts: SelectionEditReceipt[] = []
  private readonly invocations = new Map<string, Promise<SelectionEditReceipt>>()
  private listeners = new Set<() => void>()
  private disposed = false

  constructor(options: { maxQueued?: number } = {}) {
    this.capacity = options.maxQueued ?? EDIT_QUEUE_DEFAULT_CAPACITY
    if (!Number.isSafeInteger(this.capacity) || this.capacity < 1)
      throw new RangeError('Queue capacity must be a positive integer')
  }

  get isDisposed(): boolean {
    return this.disposed
  }

  enqueue(
    input: SelectionEditTaskInput,
    runner: SelectionEditRunner,
  ): Promise<SelectionEditReceipt> {
    if (this.disposed) throw new Error('Selection edit queue is disposed')
    const existing = this.invocations.get(input.invocationId)
    if (existing) return existing
    if (this.entries.length + (this.active ? 1 : 0) >= this.capacity)
      throw new EditQueueCapacityError()
    if (!input.invocationId || !input.instruction.trim())
      throw new TypeError('Edit task is incomplete')
    const copied: SelectionEditTaskInput = { ...input, scope: validateAndCopyScope(input.scope) }
    let resolve!: Deferred['resolve']
    let reject!: Deferred['reject']
    const promise = new Promise<SelectionEditReceipt>((ok, fail) => {
      resolve = ok
      reject = fail
    })
    const entry: Entry = {
      taskId: nextTaskId(),
      input: copied,
      runner,
      deferred: { promise, resolve, reject },
      controller: new AbortController(),
      state: 'queued',
      writeStarted: false,
    }
    this.entries.push(entry)
    this.invocations.set(input.invocationId, promise)
    this.emit()
    queueMicrotask(() => void this.pump())
    return promise
  }

  invalidate(identity: Pick<SelectionScope, 'documentId' | 'sessionId' | 'generation'>): void {
    const all = [...this.entries, ...(this.active ? [this.active] : [])]
    if (
      all.some(
        ({ input }) =>
          input.scope.documentId !== identity.documentId ||
          input.scope.sessionId !== identity.sessionId ||
          input.scope.generation !== identity.generation,
      )
    )
      this.cancelAll('selection changed')
  }

  cancelAll(reason = 'cancelled'): void {
    for (const entry of this.entries.splice(0)) {
      entry.state = 'settled'
      entry.controller.abort(reason)
      entry.deferred.reject(abortError(reason))
    }
    this.active?.controller.abort(reason)
    this.paused = false
    this.emit()
  }

  resume(): void {
    if (!this.paused) return
    this.paused = false
    this.emit()
    void this.pump()
  }

  snapshot(): EditQueueSnapshot {
    return {
      queued: this.entries.length,
      running: this.active !== null,
      paused: this.paused,
      receipts: [...this.receipts],
      ...(this.active ? { activeTaskId: this.active.taskId } : {}),
    }
  }

  subscribe(listener: () => void): () => void {
    this.listeners.add(listener)
    return () => this.listeners.delete(listener)
  }

  dispose(): void {
    if (this.disposed) return
    this.disposed = true
    this.cancelAll('queue disposed')
    this.listeners.clear()
  }

  private emit(): void {
    for (const listener of this.listeners) listener()
  }

  private async pump(): Promise<void> {
    if (this.active || this.paused) return
    const entry = this.entries.shift()
    if (!entry) return
    this.active = entry
    entry.state = 'running'
    this.emit()
    try {
      const receipt = await entry.runner({
        taskId: entry.taskId,
        invocationId: entry.input.invocationId,
        instruction: entry.input.instruction,
        scope: entry.input.scope,
        signal: entry.controller.signal,
        markWriteStarted: () => {
          entry.writeStarted = true
        },
      })
      // Abort is authoritative only before write dispatch. Afterwards the host
      // receipt is the source of truth and must not be relabelled cancelled.
      if (entry.controller.signal.aborted && !entry.writeStarted)
        entry.deferred.reject(abortError(String(entry.controller.signal.reason ?? 'cancelled')))
      else {
        this.receipts.push(receipt)
        if (this.receipts.length > this.capacity) this.receipts.shift()
        entry.deferred.resolve(receipt)
        if (receipt.status === 'uncertain') this.paused = true
      }
    } catch (error) {
      entry.deferred.reject(error)
    } finally {
      entry.state = 'settled'
      this.active = null
      this.emit()
      void this.pump()
    }
  }
}
