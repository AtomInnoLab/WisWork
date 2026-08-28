import { describe, expect, it, vi } from 'vitest'
import {
  SelectionEditQueue,
  SelectionScopeConflict,
  assertOperationsWithinSelectionScope,
  type SelectionEditReceipt,
  type SelectionScope,
} from '../src/renderer/ai/edit-queue'

const scope = (generation = 1): SelectionScope => ({
  documentId: 'deck-a',
  sessionId: 'session-a',
  generation,
  slides: [
    {
      slideId: 'slide-1',
      elements: [
        {
          elementId: 'shape-1',
          expectedType: 'shape',
          expectedFingerprint: `sha256:${'a'.repeat(64)}`,
        },
      ],
    },
  ],
})

const applied = (taskId: string): SelectionEditReceipt => ({
  taskId,
  invocationId: `inv-${taskId}`,
  transactionId: `tx-${taskId}`,
  status: 'applied',
})

describe('SelectionEditQueue', () => {
  it('runs different tasks FIFO and continues after an ordinary failure', async () => {
    const order: string[] = []
    const queue = new SelectionEditQueue({ maxQueued: 20 })
    const first = queue.enqueue(
      { invocationId: 'one', instruction: 'one', scope: scope() },
      async ({ taskId }) => {
        order.push('one:start')
        await Promise.resolve()
        order.push('one:end')
        throw new Error(taskId)
      },
    )
    const second = queue.enqueue(
      { invocationId: 'two', instruction: 'two', scope: scope() },
      async ({ taskId }) => {
        order.push('two')
        return applied(taskId)
      },
    )
    await expect(first).rejects.toThrow()
    await expect(second).resolves.toMatchObject({ status: 'applied' })
    expect(order).toEqual(['one:start', 'one:end', 'two'])
  })

  it('deduplicates the same invocation and enforces capacity', async () => {
    let release!: () => void
    const gate = new Promise<void>((resolve) => (release = resolve))
    const run = vi.fn(async ({ taskId }: { taskId: string }) => {
      await gate
      return applied(taskId)
    })
    const queue = new SelectionEditQueue({ maxQueued: 2 })
    const one = queue.enqueue({ invocationId: 'same', instruction: 'x', scope: scope() }, run)
    expect(queue.enqueue({ invocationId: 'same', instruction: 'x', scope: scope() }, run)).toBe(one)
    queue.enqueue({ invocationId: 'two', instruction: 'x', scope: scope() }, run)
    expect(() =>
      queue.enqueue({ invocationId: 'three', instruction: 'x', scope: scope() }, run),
    ).toThrow(/capacity/i)
    release()
    await one
    expect(run).toHaveBeenCalledTimes(2)
  })

  it('cancels queued and active work on selection generation or document changes without retargeting', async () => {
    let activeSignal!: AbortSignal
    const queue = new SelectionEditQueue()
    const active = queue.enqueue(
      { invocationId: 'one', instruction: 'x', scope: scope(1) },
      async ({ signal }) => {
        activeSignal = signal
        await new Promise((resolve) => signal.addEventListener('abort', resolve, { once: true }))
        throw new DOMException('aborted', 'AbortError')
      },
    )
    const queuedRun = vi.fn()
    const queued = queue.enqueue(
      { invocationId: 'two', instruction: 'x', scope: scope(1) },
      queuedRun,
    )
    await Promise.resolve()
    queue.invalidate({ documentId: 'deck-a', sessionId: 'session-a', generation: 2 })
    await expect(active).rejects.toMatchObject({ name: 'AbortError' })
    await expect(queued).rejects.toMatchObject({ name: 'AbortError' })
    expect(activeSignal.aborted).toBe(true)
    expect(queuedRun).not.toHaveBeenCalled()
  })

  it('preserves a receipt returned after abort during the write phase', async () => {
    let enteredWrite!: () => void
    const write = new Promise<void>((resolve) => (enteredWrite = resolve))
    const queue = new SelectionEditQueue()
    const result = queue.enqueue(
      { invocationId: 'one', instruction: 'x', scope: scope() },
      async (ctx) => {
        ctx.markWriteStarted()
        enteredWrite()
        await Promise.resolve()
        return applied(ctx.taskId)
      },
    )
    await write
    queue.cancelAll('stop')
    await expect(result).resolves.toMatchObject({ status: 'applied' })
  })

  it('pauses after uncertain and can resume remaining work', async () => {
    const queue = new SelectionEditQueue()
    const first = queue.enqueue(
      { invocationId: 'one', instruction: 'x', scope: scope() },
      async ({ taskId }) => ({
        ...applied(taskId),
        status: 'uncertain' as const,
      }),
    )
    const nextRun = vi.fn(async ({ taskId }: { taskId: string }) => applied(taskId))
    const second = queue.enqueue({ invocationId: 'two', instruction: 'x', scope: scope() }, nextRun)
    await first
    await Promise.resolve()
    expect(queue.snapshot().paused).toBe(true)
    expect(nextRun).not.toHaveBeenCalled()
    queue.resume()
    await second
    expect(nextRun).toHaveBeenCalledOnce()
    expect(queue.snapshot().receipts.map((receipt) => receipt.status)).toEqual([
      'uncertain',
      'applied',
    ])
  })
})

describe('selection scope enforcement', () => {
  it('rejects duplicate and oversized captures, including mixed slides', () => {
    expect(() =>
      new SelectionEditQueue().enqueue(
        {
          invocationId: 'dup',
          instruction: 'x',
          scope: {
            ...scope(),
            slides: [{ slideId: 'slide-1' }, { slideId: 'slide-1' }],
          },
        },
        async ({ taskId }) => applied(taskId),
      ),
    ).toThrow(/duplicate/i)
    expect(() =>
      new SelectionEditQueue().enqueue(
        {
          invocationId: 'large',
          instruction: 'x',
          scope: {
            ...scope(),
            slides: Array.from({ length: 11 }, (_, index) => ({ slideId: `slide-${index}` })),
          },
        },
        async ({ taskId }) => applied(taskId),
      ),
    ).toThrow(/1-10/)
  })

  it('rejects stale and out-of-scope targets before mutation', () => {
    expect(() =>
      assertOperationsWithinSelectionScope(scope(), [{ target: { slideId: 'slide-2' } }]),
    ).toThrow(SelectionScopeConflict)
    expect(() =>
      assertOperationsWithinSelectionScope(scope(), [
        { target: { slideId: 'slide-1', elementId: 'shape-2' } },
      ]),
    ).toThrow(SelectionScopeConflict)
    expect(() =>
      assertOperationsWithinSelectionScope(scope(), [{ target: { slideId: 'slide-1' } }]),
    ).toThrow(SelectionScopeConflict)
    expect(() =>
      assertOperationsWithinSelectionScope(scope(), [
        { target: { slideId: 'slide-1', elementId: 'shape-1', expectedType: 'picture' } },
      ]),
    ).toThrow(SelectionScopeConflict)
  })

  it('permits only an exactly captured element target', () => {
    expect(() =>
      assertOperationsWithinSelectionScope(scope(), [
        {
          target: {
            slideId: 'slide-1',
            elementId: 'shape-1',
            expectedType: 'shape',
            expectedFingerprint: `sha256:${'a'.repeat(64)}`,
          },
        },
      ]),
    ).not.toThrow()
  })
})
