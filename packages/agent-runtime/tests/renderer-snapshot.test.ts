import { describe, expect, it, vi } from 'vitest'
import { createEnhancedRendererClient } from '../src/renderer'

const mutation = { id: 'call-1', name: 'replace_blocks', input: {} }

describe('enhanced renderer mutation snapshots', () => {
  it('captures before execution and restores the private snapshot on the exact result event', async () => {
    let toolListener!: (request: any) => void
    let sessionListener!: (event: any) => void
    const order: string[] = []
    const results: any[] = []
    const bridge: any = {
      register: vi.fn(async () => undefined),
      unregister: vi.fn(async () => undefined),
      status: vi.fn(async () => ({ activeAgentRuntime: 'enhanced', documentId: 'doc' })),
      startTurn: vi.fn(async () => undefined),
      cancelTurn: vi.fn(async () => undefined),
      subscribe: (_id: string, listener: any) => {
        sessionListener = listener
        return () => undefined
      },
      onToolCall: (listener: any) => {
        toolListener = listener
        return () => undefined
      },
      toolResult: vi.fn(async (result: any) => {
        results.push(result)
        sessionListener({
          type: 'tool-executed',
          event: {
            call: mutation,
            execution: result.execution,
            snapshotBefore: result.snapshotBefore,
          },
        })
      }),
    }
    const snapshot = { type: 'doc-before' }
    const client = createEnhancedRendererClient(bridge)
    const session = client.open({
      host: 'docs',
      documentId: 'doc',
      generation: 1,
      skill: {
        id: 'docs',
        systemPrompt: '',
        tools: [{ name: 'replace_blocks', description: '', inputSchema: {} }],
        executeTool: vi.fn(() => {
          order.push('execute')
          return { output: 'ok', summary: 'ok', mutated: true }
        }),
      },
      captureSnapshot: () => {
        order.push('snapshot')
        return snapshot
      },
    })
    const events: any[] = []
    session.subscribe((event) => events.push(event))
    toolListener({ documentId: 'doc', generation: 1, call: mutation })
    await vi.waitFor(() => expect(results).toHaveLength(1))
    expect(order).toEqual(['snapshot', 'execute'])
    expect(typeof results[0].snapshotBefore).toBe('string')
    expect(events.at(-1).event.snapshotBefore).toBe(snapshot)
  })

  it('does not capture or execute a cancelled proposal and fails closed without snapshot authority', async () => {
    let toolListener!: (request: any) => void
    const captureSnapshot = vi.fn()
    const executeTool = vi.fn()
    const toolResult = vi.fn(async () => undefined)
    const bridge: any = {
      register: vi.fn(async () => undefined),
      unregister: vi.fn(async () => undefined),
      status: vi.fn(),
      startTurn: vi.fn(),
      cancelTurn: vi.fn(),
      subscribe: () => () => undefined,
      onToolCall: (listener: any) => {
        toolListener = listener
        return () => undefined
      },
      toolResult,
    }
    createEnhancedRendererClient(bridge).open({
      host: 'docs',
      documentId: 'doc',
      generation: 1,
      skill: {
        id: 'docs',
        systemPrompt: '',
        tools: [{ name: 'replace_blocks', description: '', inputSchema: {} }],
        executeTool,
      },
    })
    expect(captureSnapshot).not.toHaveBeenCalled()
    expect(executeTool).not.toHaveBeenCalled()
    toolListener({ documentId: 'doc', generation: 1, call: mutation })
    await vi.waitFor(() => expect(toolResult).toHaveBeenCalledOnce())
    expect(executeTool).not.toHaveBeenCalled()
    expect((toolResult.mock.calls as any)[0][0].execution).toMatchObject({
      output: 'snapshot_unavailable',
      isError: true,
    })
  })

  it('lets Sheets use its operation transaction and undo when no generic snapshot is supplied', async () => {
    let toolListener!: (request: any) => void
    const executeTool = vi.fn(() => ({ output: 'ok', summary: 'applied', mutated: true }))
    const toolResult = vi.fn(async () => undefined)
    const bridge: any = {
      register: vi.fn(async () => undefined),
      unregister: vi.fn(async () => undefined),
      status: vi.fn(),
      startTurn: vi.fn(),
      cancelTurn: vi.fn(),
      subscribe: () => () => undefined,
      onToolCall: (listener: any) => {
        toolListener = listener
        return () => undefined
      },
      toolResult,
    }
    createEnhancedRendererClient(bridge).open({
      host: 'sheets',
      documentId: 'sheet',
      generation: 1,
      skill: {
        id: 'sheets',
        systemPrompt: '',
        tools: [{ name: 'propose_operations', description: '', inputSchema: {} }],
        executeTool,
      },
    })
    toolListener({
      documentId: 'sheet',
      generation: 1,
      call: { id: 's1', name: 'propose_operations', input: { operations: [] } },
    })
    await vi.waitFor(() => expect(toolResult).toHaveBeenCalledOnce())
    expect(executeTool).toHaveBeenCalledOnce()
    expect((toolResult.mock.calls as any)[0][0]).not.toHaveProperty('snapshotBefore')
    expect((toolResult.mock.calls as any)[0][0].execution).toMatchObject({ mutated: true })
  })
})
