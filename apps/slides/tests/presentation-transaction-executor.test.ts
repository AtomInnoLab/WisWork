import { describe, expect, it } from 'vitest'
import type { PresentationOperation, PresentationTransaction } from '@wiswork/presentation-ops'
import {
  PresentationTransactionExecutor,
  type AtomicPresentationHost,
  type PlannedPresentationOperation,
} from '../src/main/operations/executor'

type State = { revision: string; values: Map<string, string> }
type Snapshot = { revision: string; values: Map<string, string> }

const fp = (char: string) => `sha256:${char.repeat(64)}`

function operation(clientId: string, elementId: string, text: string): PresentationOperation {
  return {
    kind: 'set_text',
    clientId,
    target: {
      slideId: 'slide-1',
      elementId,
      expectedType: 'text',
      expectedFingerprint: fp('a'),
    },
    text,
  }
}

function transaction(operations: PresentationOperation[]): PresentationTransaction {
  return {
    transactionId: 'tx-1',
    expectedDeckRevision: fp('0'),
    operations,
    mode: 'atomic',
  }
}

function fakeHost(options?: {
  throwAt?: number
  abortAt?: number
  thirdStateAt?: number
  visibilityDelay?: number
}) {
  const state: State = { revision: fp('0'), values: new Map([['shape-1', 'before']]) }
  let applyCount = 0
  let verifyCount = 0
  let historyCount = 0
  let restoreCount = 0
  let generated = 0
  const controller = new AbortController()

  const host: AtomicPresentationHost<Snapshot> = {
    readRevision: async () => state.revision,
    captureSnapshot: async () => ({ revision: state.revision, values: new Map(state.values) }),
    plan: async (_snapshot, ops, allocateId) => {
      const simulated = new Map(state.values)
      const planned: PlannedPresentationOperation[] = []
      for (const [index, op] of ops.entries()) {
        const createdId = op.kind === 'add_text_box' ? allocateId(op.clientId) : undefined
        if (op.kind === 'set_text') simulated.set(op.target.elementId!, op.text)
        if (op.kind === 'add_text_box') simulated.set(createdId!, op.text)
        planned.push({ index, operation: op, ...(createdId ? { createdId } : {}) })
      }
      return { status: 'planned', operations: planned, noOp: false }
    },
    allocateElementId: () => `created-${++generated}`,
    apply: async (planned) => {
      applyCount += 1
      const op = planned.operation
      if (op.kind === 'set_text') state.values.set(op.target.elementId!, op.text)
      if (op.kind === 'add_text_box') state.values.set(planned.createdId!, op.text)
      state.revision = fp(String(applyCount))
      if (options?.thirdStateAt === applyCount) state.revision = fp('f')
      if (options?.abortAt === applyCount) controller.abort()
      if (options?.throwAt === applyCount) throw new Error('host failure')
      return { revision: state.revision }
    },
    verify: async () => {
      verifyCount += 1
      if (verifyCount <= (options?.visibilityDelay ?? 0)) return { status: 'pending' }
      return { status: 'matched', revision: state.revision }
    },
    isAttributableRevision: async (revision) => revision !== fp('f'),
    restoreIfCurrent: async (expected, snapshot) => {
      if (state.revision !== expected) return 'different'
      restoreCount += 1
      state.revision = snapshot.revision
      state.values = new Map(snapshot.values)
      return 'restored'
    },
    publishHistory: async () => {
      historyCount += 1
      return true
    },
  }
  return {
    host,
    state,
    controller,
    counts: () => ({ applyCount, verifyCount, historyCount, restoreCount, generated }),
  }
}

describe('PresentationTransactionExecutor', () => {
  it('plans dependent operations sequentially and creates one history entry', async () => {
    const fixture = fakeHost()
    const receipt = await new PresentationTransactionExecutor(fixture.host).execute(
      transaction([operation('a', 'shape-1', 'one'), operation('b', 'shape-1', 'two')]),
    )
    expect(receipt).toMatchObject({ status: 'applied', operationCount: 2 })
    expect(fixture.state.values.get('shape-1')).toBe('two')
    expect(fixture.counts().historyCount).toBe(1)
  })

  it('restores only an attributable intermediate state after a mid-operation throw', async () => {
    const fixture = fakeHost({ throwAt: 2 })
    const receipt = await new PresentationTransactionExecutor(fixture.host).execute(
      transaction([operation('a', 'shape-1', 'one'), operation('b', 'shape-1', 'two')]),
    )
    expect(receipt.status).toBe('unchanged')
    expect(fixture.state.values.get('shape-1')).toBe('before')
    expect(fixture.counts()).toMatchObject({ restoreCount: 1, historyCount: 0 })
  })

  it('never restores a concurrent third state', async () => {
    const fixture = fakeHost({ throwAt: 2, thirdStateAt: 2 })
    const receipt = await new PresentationTransactionExecutor(fixture.host).execute(
      transaction([operation('a', 'shape-1', 'one'), operation('b', 'shape-1', 'two')]),
    )
    expect(receipt.status).toBe('uncertain')
    expect(fixture.counts().restoreCount).toBe(0)
  })

  it('handles abort before apply as unchanged', async () => {
    const fixture = fakeHost()
    fixture.controller.abort()
    const receipt = await new PresentationTransactionExecutor(fixture.host).execute(
      transaction([operation('a', 'shape-1', 'one')]),
      fixture.controller.signal,
    )
    expect(receipt.status).toBe('unchanged')
    expect(fixture.counts().applyCount).toBe(0)
  })

  it('restores an attributable state when aborted during apply', async () => {
    const fixture = fakeHost({ abortAt: 1 })
    const receipt = await new PresentationTransactionExecutor(fixture.host).execute(
      transaction([operation('a', 'shape-1', 'one'), operation('b', 'shape-1', 'two')]),
      fixture.controller.signal,
    )
    expect(receipt.status).toBe('unchanged')
    expect(fixture.counts().restoreCount).toBe(1)
  })

  it('verifies a completed write even when cancellation arrives after apply', async () => {
    const fixture = fakeHost({ abortAt: 1 })
    const receipt = await new PresentationTransactionExecutor(fixture.host).execute(
      transaction([operation('a', 'shape-1', 'one')]),
      fixture.controller.signal,
    )
    expect(receipt.status).toBe('applied')
    expect(fixture.counts().historyCount).toBe(1)
  })

  it('bounds delayed-visibility verification', async () => {
    const fixture = fakeHost({ visibilityDelay: 2 })
    const receipt = await new PresentationTransactionExecutor(fixture.host, {
      verifyAttempts: 3,
      verifyDelayMs: 0,
    }).execute(transaction([operation('a', 'shape-1', 'one')]))
    expect(receipt.status).toBe('applied')
    expect(fixture.counts().verifyCount).toBe(3)
  })

  it('reports uncertain when the final history CAS observes a third mutation', async () => {
    const fixture = fakeHost()
    fixture.host.publishHistory = async () => false
    const receipt = await new PresentationTransactionExecutor(fixture.host).execute(
      transaction([operation('a', 'shape-1', 'one')]),
    )
    expect(receipt).toMatchObject({ status: 'uncertain', code: 'write_state_uncertain' })
    expect(fixture.counts().historyCount).toBe(0)
  })

  it('rejects a stale deck before planning or applying', async () => {
    const fixture = fakeHost()
    fixture.state.revision = fp('9')
    const receipt = await new PresentationTransactionExecutor(fixture.host).execute(
      transaction([operation('a', 'shape-1', 'one')]),
    )
    expect(receipt).toMatchObject({ status: 'conflict', code: 'target_stale' })
    expect(fixture.counts().applyCount).toBe(0)
  })

  it('rechecks the authoritative revision after asynchronous planning', async () => {
    const fixture = fakeHost()
    const originalPlan = fixture.host.plan
    fixture.host.plan = async (...args) => {
      const plan = await originalPlan(...args)
      fixture.state.revision = fp('8')
      return plan
    }
    const receipt = await new PresentationTransactionExecutor(fixture.host).execute(
      transaction([operation('a', 'shape-1', 'one')]),
    )
    expect(receipt).toMatchObject({ status: 'conflict', code: 'target_stale' })
    expect(fixture.counts().applyCount).toBe(0)
  })

  it('does not restore when the host cannot provide compare-and-swap', async () => {
    const fixture = fakeHost({ throwAt: 1 })
    fixture.host.restoreIfCurrent = async () => 'unsupported'
    const receipt = await new PresentationTransactionExecutor(fixture.host).execute(
      transaction([operation('a', 'shape-1', 'one')]),
    )
    expect(receipt.status).toBe('uncertain')
    expect(fixture.counts().historyCount).toBe(0)
  })

  it.each(['verify', 'post_write_read', 'recover_read', 'attribution', 'history_publish'] as const)(
    'bounds %s exceptions after the write phase as uncertain',
    async (phase) => {
      const fixture = fakeHost(
        phase === 'recover_read' || phase === 'attribution' ? { throwAt: 1 } : {},
      )
      if (phase === 'verify') fixture.host.verify = async () => Promise.reject(new Error('secret'))
      if (phase === 'post_write_read') {
        const originalRead = fixture.host.readRevision
        let reads = 0
        fixture.host.verify = async () => ({ status: 'mismatch' })
        fixture.host.readRevision = async () => {
          reads += 1
          if (reads === 3) throw new Error('secret')
          return originalRead()
        }
      }
      if (phase === 'recover_read') {
        const originalRead = fixture.host.readRevision
        let reads = 0
        fixture.host.readRevision = async () => {
          reads += 1
          if (reads === 3) throw new Error('secret')
          return originalRead()
        }
      }
      if (phase === 'attribution') {
        fixture.host.isAttributableRevision = async () => Promise.reject(new Error('secret'))
      }
      if (phase === 'history_publish') {
        fixture.host.publishHistory = async () => Promise.reject(new Error('secret'))
      }
      const receipt = await new PresentationTransactionExecutor(fixture.host).execute(
        transaction([operation('a', 'shape-1', 'one')]),
      )
      expect(receipt).toEqual({
        status: 'uncertain',
        transactionId: 'tx-1',
        code: 'write_state_uncertain',
      })
      expect(JSON.stringify(receipt)).not.toContain('secret')
    },
  )

  it.each(['initial_read', 'snapshot', 'plan', 'prewrite_read'] as const)(
    'bounds %s exceptions before the write phase as unchanged',
    async (phase) => {
      const fixture = fakeHost()
      if (phase === 'initial_read')
        fixture.host.readRevision = async () => Promise.reject(new Error('secret'))
      if (phase === 'snapshot')
        fixture.host.captureSnapshot = async () => Promise.reject(new Error('secret'))
      if (phase === 'plan') fixture.host.plan = async () => Promise.reject(new Error('secret'))
      if (phase === 'prewrite_read') {
        const originalRead = fixture.host.readRevision
        let reads = 0
        fixture.host.readRevision = async () => {
          reads += 1
          if (reads === 2) throw new Error('secret')
          return originalRead()
        }
      }
      const receipt = await new PresentationTransactionExecutor(fixture.host).execute(
        transaction([operation('a', 'shape-1', 'one')]),
      )
      expect(receipt).toEqual({
        status: 'unchanged',
        transactionId: 'tx-1',
        code: 'write_not_applied',
        operationCount: 1,
      })
      expect(fixture.counts()).toMatchObject({ applyCount: 0, historyCount: 0 })
      expect(JSON.stringify(receipt)).not.toContain('secret')
    },
  )

  it('returns unchanged for a planned no-op', async () => {
    const fixture = fakeHost()
    fixture.host.plan = async () => ({ status: 'planned', operations: [], noOp: true })
    const receipt = await new PresentationTransactionExecutor(fixture.host).execute(
      transaction([operation('a', 'shape-1', 'before')]),
    )
    expect(receipt).toMatchObject({ status: 'unchanged', code: 'operation_noop' })
    expect(fixture.counts().historyCount).toBe(0)
  })

  it('reports generated IDs and retries the same transaction idempotently', async () => {
    const fixture = fakeHost()
    const tx = transaction([
      {
        kind: 'add_text_box',
        clientId: 'new',
        slideId: 'slide-1',
        text: 'hello',
        geometry: { x: 1, y: 2, width: 3, height: 4 },
      },
    ])
    const executor = new PresentationTransactionExecutor(fixture.host)
    const first = await executor.execute(tx)
    const second = await executor.execute(tx)
    expect(first).toMatchObject({ status: 'applied', createdIds: ['created-1'] })
    expect(second).toEqual(first)
    expect(fixture.counts()).toMatchObject({ applyCount: 1, historyCount: 1, generated: 1 })
  })

  it('singleflights concurrent duplicate ids before acquiring one write lease', async () => {
    const fixture = fakeHost()
    const originalApply = fixture.host.apply
    let resume!: () => void
    let started!: () => void
    const paused = new Promise<void>((resolve) => {
      resume = resolve
    })
    const applying = new Promise<void>((resolve) => {
      started = resolve
    })
    fixture.host.apply = async (planned, signal) => {
      started()
      await paused
      return originalApply(planned, signal)
    }
    let leases = 0
    let releases = 0
    const executor = new PresentationTransactionExecutor(fixture.host, {
      acquireWriteLease: () => {
        leases += 1
        return () => {
          releases += 1
        }
      },
    })
    const tx = transaction([operation('a', 'shape-1', 'one')])
    const first = executor.execute(tx)
    await applying
    const duplicate = executor.execute(structuredClone(tx))

    expect(duplicate).toBe(first)
    expect(leases).toBe(1)
    resume()
    const [firstReceipt, duplicateReceipt] = await Promise.all([first, duplicate])
    expect(duplicateReceipt).toEqual(firstReceipt)
    expect(duplicateReceipt).toBe(firstReceipt)
    expect({ leases, releases }).toEqual({ leases: 1, releases: 1 })
    expect(fixture.counts()).toMatchObject({ applyCount: 1, historyCount: 1 })
  })

  it('returns an exact cached receipt while a different transaction holds the write lease', async () => {
    const fixture = fakeHost()
    const executor = new PresentationTransactionExecutor(fixture.host)
    const cachedTx = transaction([operation('a', 'shape-1', 'one')])
    const cachedReceipt = await executor.execute(cachedTx)

    let resume!: () => void
    let started!: () => void
    const paused = new Promise<void>((resolve) => {
      resume = resolve
    })
    const applying = new Promise<void>((resolve) => {
      started = resolve
    })
    const originalApply = fixture.host.apply
    fixture.host.apply = async (planned, signal) => {
      started()
      await paused
      return originalApply(planned, signal)
    }
    const other = {
      ...transaction([operation('b', 'shape-1', 'two')]),
      transactionId: 'tx-2',
      expectedDeckRevision: fixture.state.revision,
    }
    const busy = executor.execute(other)
    await applying
    const retry = await executor.execute(structuredClone(cachedTx))

    expect(retry).toEqual(cachedReceipt)
    expect(retry).toBe(cachedReceipt)
    resume()
    await busy
  })

  it('fails closed at ledger capacity without forgetting an applied transaction', async () => {
    const fixture = fakeHost()
    const executor = new PresentationTransactionExecutor(fixture.host, {
      maxCachedReceipts: 1,
    })
    const first = transaction([operation('a', 'shape-1', 'one')])
    const firstReceipt = await executor.execute(first)
    const second = {
      ...transaction([operation('b', 'shape-1', 'two')]),
      transactionId: 'tx-2',
      expectedDeckRevision: fixture.state.revision,
    }
    expect(await executor.execute(second)).toMatchObject({
      status: 'unchanged',
      code: 'write_not_applied',
    })
    expect(await executor.execute(first)).toEqual(firstReceipt)
    expect(fixture.counts().applyCount).toBe(1)
  })
})
