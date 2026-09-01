import { describe, expect, it, vi } from 'vitest'
import { runEnhancedGolden } from '../src/production-golden'

describe('production Enhanced golden orchestration', () => {
  it('runs registration, turn, production skill, confirmation, receipt/readback and rollback in order', async () => {
    const order: string[] = []
    let registered: any
    let emitToolCall: ((request: any) => void) | undefined
    let emitEvent: ((event: any) => void) | undefined
    const bridge: any = {
      status: vi.fn(async () => ({ activeAgentRuntime: 'enhanced', documentId: 'deck-1' })),
      register: vi.fn(async (input) => {
        registered = input
        order.push('register')
      }),
      unregister: vi.fn(async () => undefined),
      subscribe: vi.fn((_documentId, listener) => {
        emitEvent = listener
        return () => undefined
      }),
      onToolCall: vi.fn((listener) => {
        emitToolCall = listener
        return () => undefined
      }),
      startTurn: vi.fn(async () => {
        order.push('turn')
        emitToolCall?.({
          documentId: 'deck-1',
          generation: 4,
          call: { id: 'call-1', name: 'set_element_text', input: { text: 'after' } },
        })
      }),
      toolResult: vi.fn(async (result) => {
        order.push('tool-result')
        emitEvent?.({
          type: 'tool-executed',
          event: {
            call: { id: result.callId, name: 'set_element_text', input: { text: 'after' } },
            execution: result.execution,
          },
        })
      }),
      cancelTurn: vi.fn(async () => undefined),
    }
    const result = await runEnhancedGolden('slides', {
      documentId: 'deck-1',
      generation: 4,
      instruction: 'change the title',
      bridge,
      skill: {
        id: 'real-slides-skill',
        systemPrompt: 'bounded',
        tools: [{ name: 'set_element_text', description: 'edit', inputSchema: { type: 'object' } }],
        executeTool: vi.fn(async () => {
          order.push('skill')
          return { output: 'proposal-1', summary: 'pending', mutated: false }
        }),
      },
      confirm: async () => {
        order.push('confirm')
        return { mutationReceiptId: 'receipt-1' }
      },
      readback: async () => {
        order.push('readback')
        return { status: 'verified' as const }
      },
      rollback: async () => {
        order.push('rollback')
        return { status: 'restored' as const }
      },
    })

    expect(registered).toMatchObject({
      host: 'slides',
      documentId: 'deck-1',
      generation: 4,
      mutatingTools: ['set_element_text'],
    })
    expect(order).toEqual([
      'register',
      'turn',
      'skill',
      'tool-result',
      'confirm',
      'readback',
      'rollback',
    ])
    expect(result).toEqual({
      host: 'slides',
      mutationReceiptId: 'receipt-1',
      verification: { status: 'verified' },
      rollback: { status: 'restored' },
    })
  })
})
