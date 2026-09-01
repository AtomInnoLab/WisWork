import { isAbsolute } from 'node:path'
import type { MessagesRequest } from '@wiswork/codex-bridge'
import { describe, expect, it, vi } from 'vitest'
import {
  createProductionCodexBootstrap,
  startBestEffortCodexInterrupt,
} from '../src/main/codex-engine'

const executable = process.env.WISWORK_CODEX_INTEGRATION_EXECUTABLE
const realIt = executable && isAbsolute(executable) ? it : it.skip

it('detaches an unresponsive interrupt and bounds its lifetime', async () => {
  vi.useFakeTimers()
  const interrupt = vi.fn(() => new Promise<never>(() => undefined))
  startBestEffortCodexInterrupt(interrupt)
  await Promise.resolve()
  expect(interrupt).toHaveBeenCalledOnce()
  await vi.advanceTimersByTimeAsync(1_999)
  expect(vi.getTimerCount()).toBe(1)
  await vi.advanceTimersByTimeAsync(1)
  expect(vi.getTimerCount()).toBe(0)
  vi.useRealTimers()
})

function finalResponse(): Response {
  return new Response(
    [
      'event: message_start\ndata: {"type":"message_start","message":{"id":"msg_1","model":"openai/gpt-5.6-sol","usage":{"input_tokens":1}}}\n\n',
      'event: content_block_start\ndata: {"type":"content_block_start","index":0,"content_block":{"type":"text","text":""}}\n\n',
      'event: content_block_delta\ndata: {"type":"content_block_delta","index":0,"delta":{"type":"text_delta","text":"done"}}\n\n',
      'event: content_block_stop\ndata: {"type":"content_block_stop","index":0}\n\n',
      'event: message_delta\ndata: {"type":"message_delta","delta":{"stop_reason":"end_turn"},"usage":{"output_tokens":1}}\n\n',
      'event: message_stop\ndata: {"type":"message_stop"}\n\n',
    ].join(''),
    { status: 200, headers: { 'content-type': 'text/event-stream' } },
  )
}

describe('real 0.147 production engine bridge', () => {
  realIt(
    'binds real turn metadata to fake WisUsage and cleans up',
    async () => {
      const diagnostics: string[] = []
      const upstream = vi.fn(async (request: MessagesRequest) => {
        expect(request.tools?.map((tool) => tool.name)).toEqual(['exec'])
        return finalResponse()
      })
      const engine = await createProductionCodexBootstrap({
        fetchWithAuth: upstream,
        diagnostics: (code) => diagnostics.push(code),
      }).start({ executablePath: executable!, onCrash: vi.fn() })
      const session = {
        identity: {
          ownerId: 'owner',
          host: 'docs',
          documentId: 'doc',
          sessionId: 'session',
          generation: 1,
        },
        credentials: { sessionId: 'session', secret: 'secret' },
        callTool: vi.fn(async () => ({ output: 'read', summary: 'read' })),
        close: vi.fn(),
      } as any
      const events: unknown[] = []
      engine.registerDocument!({
        ownerId: 'owner',
        documentId: 'doc',
        host: 'docs',
        generation: 1,
        session,
        onEvent: (event) => events.push(event),
      })
      try {
        await engine
          .startTurn({ documentId: 'doc', host: 'docs', generation: 1, text: 'read' })
          .catch((error) => {
            throw new Error(`engine_failed:${diagnostics.join(',')}`, { cause: error })
          })
        expect(upstream).toHaveBeenCalledOnce()
      } finally {
        await engine.close()
      }
      expect(session.close).toHaveBeenCalled()
      expect(events).toEqual(
        expect.arrayContaining([
          expect.objectContaining({ type: 'text', text: 'done' }),
          expect.objectContaining({ type: 'terminal', status: 'completed' }),
        ]),
      )
      expect(diagnostics).toContain('gateway_tools_list')
    },
    20_000,
  )

  realIt(
    'settles cancellation without waiting for a provider terminal notification',
    async () => {
      const upstream = vi.fn(
        async () =>
          new Response(new ReadableStream({ start() {} }), {
            status: 200,
            headers: { 'content-type': 'text/event-stream' },
          }),
      )
      const engine = await createProductionCodexBootstrap({ fetchWithAuth: upstream }).start({
        executablePath: executable!,
        onCrash: vi.fn(),
      })
      const events: unknown[] = []
      const session = {
        identity: {
          ownerId: 'owner',
          host: 'docs',
          documentId: 'cancel-doc',
          sessionId: 'cancel-session',
          generation: 1,
        },
        credentials: { sessionId: 'cancel-session', secret: 'cancel-secret' },
        callTool: vi.fn(),
        cancelAll: vi.fn(() => 0),
        close: vi.fn(),
      } as any
      engine.registerDocument!({
        ownerId: 'owner',
        documentId: 'cancel-doc',
        host: 'docs',
        generation: 1,
        session,
        onEvent: (event) => events.push(event),
      })
      try {
        const running = engine.startTurn({
          documentId: 'cancel-doc',
          host: 'docs',
          generation: 1,
          text: 'wait',
        })
        await vi.waitFor(() => expect(upstream).toHaveBeenCalled(), { timeout: 10_000 })
        await engine.cancelTurn('cancel-doc')
        await expect(running).resolves.toBeUndefined()
        expect(session.cancelAll).toHaveBeenCalled()
        expect(events).toContainEqual({ type: 'terminal', status: 'cancelled' })
      } finally {
        await engine.close()
      }
    },
    20_000,
  )
})
