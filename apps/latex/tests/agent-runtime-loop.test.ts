// @vitest-environment jsdom
import { describe, expect, it, vi } from 'vitest'
import { createLatexRuntimeLoop } from '../src/renderer/ai/runtime-loop'

describe('LaTeX runtime selector', () => {
  it('uses Enhanced without replaying through Standard', async () => {
    const stream = vi.fn()
    let documentId: string | null = null
    const api: any = {
      status: vi.fn(async () => ({ activeAgentRuntime: 'enhanced', documentId })),
      register: vi.fn(async (input: any) => {
        documentId = input.documentId
      }),
      unregister: vi.fn(async () => undefined),
      startTurn: vi.fn(async () => undefined),
      cancelTurn: vi.fn(async () => undefined),
      toolResult: vi.fn(async () => undefined),
      onEvent: vi.fn(() => () => undefined),
      onToolCall: vi.fn(() => () => undefined),
    }
    const loop = createLatexRuntimeLoop(
      {
        transport: { stream },
        skill: {
          id: 'latex',
          systemPrompt: 'bounded',
          tools: [
            { name: 'read_project_text', description: 'read', inputSchema: { type: 'object' } },
          ],
          executeTool: vi.fn(async () => ({ output: 'ok', summary: 'read' })),
        },
      },
      api,
    )
    await new Promise((resolve) => setTimeout(resolve, 0))
    loop.run('read')
    await new Promise((resolve) => setTimeout(resolve, 0))
    expect(api.register).toHaveBeenCalledOnce()
    expect(api.startTurn).toHaveBeenCalledOnce()
    expect(stream).not.toHaveBeenCalled()
    loop.dispose()
  })
})
