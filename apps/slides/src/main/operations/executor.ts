import {
  canonicalizeSemanticValue,
  fingerprintSemanticValue,
  parsePresentationTransaction,
  type PresentationOperation,
  type PresentationReceipt,
  type PresentationTransaction,
} from '@wiswork/presentation-ops'
import {
  validatePlannedOperations,
  type PlannedPresentationOperation,
  type PresentationPlan,
} from './planner'
import { conflictReceipt, unchangedReceipt, uncertainReceipt } from './receipts'

export type { PlannedPresentationOperation } from './planner'

export interface AtomicPresentationHost<Snapshot> {
  readRevision(): Promise<string>
  captureSnapshot(): Promise<Snapshot>
  plan(
    snapshot: Snapshot,
    operations: readonly PresentationOperation[],
    allocateId: (clientId: string) => string,
  ): Promise<PresentationPlan>
  allocateElementId(): string
  apply(planned: PlannedPresentationOperation, signal?: AbortSignal): Promise<{ revision: string }>
  verify(
    operations: readonly PlannedPresentationOperation[],
  ): Promise<
    { status: 'matched'; revision: string } | { status: 'pending' } | { status: 'mismatch' }
  >
  /** Must prove the current state is an exact transaction prefix/final state. */
  isAttributableRevision(
    revision: string,
    operations: readonly PlannedPresentationOperation[],
    appliedCount: number,
  ): Promise<boolean>
  /** Compare-and-restore. Implementations without atomic CAS must return `unsupported`. */
  restoreIfCurrent(
    expectedCurrentRevision: string,
    snapshot: Snapshot,
  ): Promise<'restored' | 'different' | 'unsupported' | 'failed'>
  /** Atomically publish only while the host snapshot lease is still current. */
  publishHistory(snapshot: Snapshot): Promise<boolean>
}

export interface PresentationTransactionExecutorOptions {
  verifyAttempts?: number
  verifyDelayMs?: number
  maxCachedReceipts?: number
  acquireWriteLease?: () => (() => void) | null
  maxQueuedTransactions?: number
  fingerprintTransaction?: (transaction: PresentationTransaction) => Promise<string>
  validateScopeGuard?: (guard: PresentationScopeGuard) => boolean
}

export interface PresentationScopeGuard {
  documentId: string
  sessionId: string
  generation: number
}

interface CachedReceipt {
  digest: string
  signature: string
  receipt: PresentationReceipt
}

export const MAX_QUEUED_PRESENTATION_TRANSACTIONS = 64

const delay = (milliseconds: number): Promise<void> =>
  milliseconds <= 0
    ? Promise.resolve()
    : new Promise((resolve) => setTimeout(resolve, milliseconds))

const synchronousTransactionSignature = (transaction: PresentationTransaction): string =>
  canonicalizeSemanticValue(transaction)

export class PresentationTransactionExecutor<Snapshot> {
  private readonly verifyAttempts: number
  private readonly verifyDelayMs: number
  private readonly maxCachedReceipts: number
  private readonly acquireWriteLease: (() => (() => void) | null) | undefined
  private readonly maxQueuedTransactions: number
  private readonly fingerprintTransaction: (transaction: PresentationTransaction) => Promise<string>
  private readonly validateScopeGuard: ((guard: PresentationScopeGuard) => boolean) | undefined
  private readonly receipts = new Map<string, CachedReceipt>()
  private readonly receiptScopeGuards = new Map<string, string>()
  private readonly inFlight = new Map<
    string,
    { signature: string; promise: Promise<PresentationReceipt> }
  >()
  private active = Promise.resolve()

  constructor(
    private readonly host: AtomicPresentationHost<Snapshot>,
    options: PresentationTransactionExecutorOptions = {},
  ) {
    this.verifyAttempts = Math.max(1, Math.min(10, options.verifyAttempts ?? 3))
    this.verifyDelayMs = Math.max(0, Math.min(1_000, options.verifyDelayMs ?? 50))
    this.maxCachedReceipts = Math.max(1, Math.min(10_000, options.maxCachedReceipts ?? 10_000))
    this.acquireWriteLease = options.acquireWriteLease
    this.maxQueuedTransactions = Math.max(
      1,
      Math.min(
        MAX_QUEUED_PRESENTATION_TRANSACTIONS,
        options.maxQueuedTransactions ?? MAX_QUEUED_PRESENTATION_TRANSACTIONS,
      ),
    )
    this.fingerprintTransaction = options.fingerprintTransaction ?? fingerprintSemanticValue
    this.validateScopeGuard = options.validateScopeGuard
  }

  execute(
    input: PresentationTransaction,
    signal?: AbortSignal,
    scopeGuard?: PresentationScopeGuard,
  ): Promise<PresentationReceipt> {
    const transaction = parsePresentationTransaction(input)
    const signature = scopeGuard
      ? canonicalizeSemanticValue({ transaction, scopeGuard })
      : synchronousTransactionSignature(transaction)
    const existing = this.inFlight.get(transaction.transactionId)
    if (existing) {
      if (existing.signature === signature) return existing.promise
      return Promise.resolve({
        status: 'conflict',
        transactionId: transaction.transactionId,
        code: 'target_stale',
      })
    }
    const cached = scopeGuard ? undefined : this.receipts.get(transaction.transactionId)
    if (cached) {
      if (cached.signature === signature) return Promise.resolve(cached.receipt)
      return Promise.resolve({
        status: 'conflict',
        transactionId: transaction.transactionId,
        code: 'target_stale',
      })
    }
    if (this.inFlight.size >= this.maxQueuedTransactions) {
      return Promise.resolve(unchangedReceipt(transaction, 'write_not_applied'))
    }
    // Occupy the FIFO slot synchronously. Digesting must happen only after the
    // preceding distinct transaction settles, otherwise a faster later digest
    // could overtake an earlier request.
    const run = this.active.then(() => this.executeQueued(transaction, signal, scopeGuard))
    this.active = run.then(
      () => undefined,
      () => undefined,
    )
    this.inFlight.set(transaction.transactionId, { signature, promise: run })
    const cleanup = () => {
      if (this.inFlight.get(transaction.transactionId)?.promise === run) {
        this.inFlight.delete(transaction.transactionId)
      }
    }
    void run.then(cleanup, cleanup)
    return run
  }

  private async executeQueued(
    transaction: PresentationTransaction,
    signal?: AbortSignal,
    scopeGuard?: PresentationScopeGuard,
  ): Promise<PresentationReceipt> {
    let digest: string
    try {
      digest = await this.fingerprintTransaction(transaction)
    } catch {
      return unchangedReceipt(transaction, 'write_not_applied')
    }
    const cached = scopeGuard ? undefined : this.receipts.get(transaction.transactionId)
    if (cached) {
      if (cached.digest === digest) return cached.receipt
      return {
        status: 'conflict',
        transactionId: transaction.transactionId,
        code: 'target_stale',
      }
    }
    return this.executeExclusive(transaction, digest, signal, scopeGuard)
  }

  private async executeExclusive(
    transaction: PresentationTransaction,
    digest: string,
    signal?: AbortSignal,
    scopeGuard?: PresentationScopeGuard,
  ): Promise<PresentationReceipt> {
    // Re-check after waiting behind a different transaction: it may have filled
    // the bounded ledger while this transaction was queued.
    const cached = scopeGuard ? undefined : this.receipts.get(transaction.transactionId)
    if (cached) {
      if (cached.digest === digest) return cached.receipt
      return {
        status: 'conflict',
        transactionId: transaction.transactionId,
        code: 'target_stale',
      }
    }
    if (!scopeGuard && this.receipts.size >= this.maxCachedReceipts) {
      // Never forget an accepted transaction id: once the bounded ledger is
      // full, fail closed for new ids instead of making old writes replayable.
      return unchangedReceipt(transaction, 'write_not_applied')
    }

    const releaseLease = this.acquireWriteLease?.()
    if (this.acquireWriteLease && !releaseLease) {
      return this.cache(transaction, digest, unchangedReceipt(transaction, 'write_not_applied'))
    }

    try {
      if (scopeGuard && (!this.validateScopeGuard || !this.validateScopeGuard(scopeGuard))) {
        return {
          status: 'conflict',
          transactionId: transaction.transactionId,
          code: 'target_stale',
        }
      }
      if (scopeGuard) {
        const guardSignature = canonicalizeSemanticValue(scopeGuard)
        const cached = this.receipts.get(transaction.transactionId)
        if (cached) {
          if (
            cached.digest !== digest ||
            this.receiptScopeGuards.get(transaction.transactionId) !== guardSignature
          ) {
            return {
              status: 'conflict',
              transactionId: transaction.transactionId,
              code: 'target_stale',
            }
          }
          return cached.receipt
        }
        if (this.receipts.size >= this.maxCachedReceipts) {
          return unchangedReceipt(transaction, 'write_not_applied')
        }
        const receipt = await this.executeWithLease(transaction, digest, signal)
        if (this.receipts.has(transaction.transactionId)) {
          this.receiptScopeGuards.set(transaction.transactionId, guardSignature)
        }
        return receipt
      }
      return await this.executeWithLease(transaction, digest, signal)
    } finally {
      releaseLease?.()
    }
  }

  private async executeWithLease(
    transaction: PresentationTransaction,
    digest: string,
    signal?: AbortSignal,
  ): Promise<PresentationReceipt> {
    let writeStarted = false
    try {
      if (signal?.aborted)
        return this.cache(transaction, digest, unchangedReceipt(transaction, 'write_not_applied'))
      const initialRevision = await this.host.readRevision()
      if (initialRevision !== transaction.expectedDeckRevision) {
        return this.cache(transaction, digest, {
          status: 'conflict',
          transactionId: transaction.transactionId,
          code: 'target_stale',
        })
      }
      const snapshot = await this.host.captureSnapshot()
      if (signal?.aborted)
        return this.cache(transaction, digest, unchangedReceipt(transaction, 'write_not_applied'))

      const allocated = new Map<string, string>()
      const plan = await this.host.plan(snapshot, transaction.operations, (clientId) => {
        const existing = allocated.get(clientId)
        if (existing) return existing
        const created = this.host.allocateElementId()
        allocated.set(clientId, created)
        return created
      })
      if (plan.status === 'conflict')
        return this.cache(transaction, digest, conflictReceipt(transaction, plan))
      validatePlannedOperations(plan.operations)
      if (plan.noOp)
        return this.cache(transaction, digest, unchangedReceipt(transaction, 'operation_noop'))
      if (signal?.aborted)
        return this.cache(transaction, digest, unchangedReceipt(transaction, 'write_not_applied'))
      // Planning may await fingerprints. Refuse to enter the write phase if any
      // ordinary editor mutation raced with the authoritative snapshot.
      if ((await this.host.readRevision()) !== initialRevision) {
        return this.cache(transaction, digest, {
          status: 'conflict',
          transactionId: transaction.transactionId,
          code: 'target_stale',
        })
      }

      writeStarted = true
      let appliedCount = 0
      let lastRevision = initialRevision
      try {
        for (const planned of plan.operations) {
          if (signal?.aborted) {
            return await this.recover(transaction, digest, snapshot, plan.operations, appliedCount)
          }
          const result = await this.host.apply(planned, signal)
          appliedCount += 1
          lastRevision = result.revision
          // Cancellation after the final apply is a completed-write verification case.
          if (signal?.aborted && appliedCount < plan.operations.length) {
            return await this.recover(transaction, digest, snapshot, plan.operations, appliedCount)
          }
        }
      } catch {
        return await this.recover(transaction, digest, snapshot, plan.operations, appliedCount)
      }

      for (let attempt = 0; attempt < this.verifyAttempts; attempt += 1) {
        const verified = await this.host.verify(plan.operations)
        if (verified.status === 'matched') {
          if (!(await this.host.publishHistory(snapshot))) {
            return this.cache(transaction, digest, uncertainReceipt(transaction))
          }
          return this.cache(transaction, digest, {
            status: 'applied',
            transactionId: transaction.transactionId,
            resultingDeckRevision: verified.revision,
            operationCount: transaction.operations.length,
            ...(allocated.size ? { createdIds: [...allocated.values()] } : {}),
            ...(allocated.size
              ? {
                  createdTargets: [...allocated].map(([clientId, elementId]) => ({
                    clientId,
                    elementId,
                  })),
                }
              : {}),
          })
        }
        if (verified.status === 'mismatch') break
        if (attempt + 1 < this.verifyAttempts) await delay(this.verifyDelayMs)
      }

      const current = await this.host.readRevision()
      if (current === initialRevision) {
        return this.cache(transaction, digest, unchangedReceipt(transaction, 'write_not_applied'))
      }
      if (current !== lastRevision) {
        return this.cache(transaction, digest, uncertainReceipt(transaction))
      }
      return await this.recover(transaction, digest, snapshot, plan.operations, appliedCount)
    } catch {
      // Before entering the write phase the deck is provably untouched. After
      // that boundary, never leak host details or claim a clean failure.
      return this.cache(
        transaction,
        digest,
        writeStarted
          ? uncertainReceipt(transaction)
          : unchangedReceipt(transaction, 'write_not_applied'),
      )
    }
  }

  private async recover(
    transaction: PresentationTransaction,
    digest: string,
    snapshot: Snapshot,
    operations: readonly PlannedPresentationOperation[],
    appliedCount: number,
  ): Promise<PresentationReceipt> {
    const current = await this.host.readRevision()
    const attributable = await this.host.isAttributableRevision(current, operations, appliedCount)
    const failureIndex = Math.min(appliedCount, transaction.operations.length - 1)
    if (!attributable)
      return this.cache(transaction, digest, uncertainReceipt(transaction, failureIndex))
    const restored = await this.host.restoreIfCurrent(current, snapshot)
    if (restored === 'restored') {
      return this.cache(transaction, digest, unchangedReceipt(transaction, 'write_not_applied'))
    }
    return this.cache(transaction, digest, uncertainReceipt(transaction, failureIndex))
  }

  private cache(
    transaction: PresentationTransaction,
    digest: string,
    receipt: PresentationReceipt,
  ): PresentationReceipt {
    this.receipts.set(transaction.transactionId, {
      digest,
      signature: synchronousTransactionSignature(transaction),
      receipt,
    })
    return receipt
  }
}
