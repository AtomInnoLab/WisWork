import { isAbsolute } from 'node:path'
import type { MessagesRequest } from '@wiswork/codex-bridge'
import { describe, expect, it, vi } from 'vitest'
import { createProductionCodexBootstrap } from '../src/main/codex-engine'

const executable = process.env.WISWORK_CODEX_INTEGRATION_EXECUTABLE
const realIt = executable && isAbsolute(executable) ? it : it.skip

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
      const upstream = vi.fn(async (_request: MessagesRequest) => finalResponse())
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
      engine.registerDocument!({
        ownerId: 'owner',
        documentId: 'doc',
        host: 'docs',
        generation: 1,
        session,
      })
      try {
        await engine.startTurn({ documentId: 'doc', host: 'docs', generation: 1, text: 'read' })
        await vi.waitFor(
          () => expect(upstream, JSON.stringify(diagnostics)).toHaveBeenCalledOnce(),
          { timeout: 10_000 },
        )
      } finally {
        await engine.close()
      }
      expect(session.close).toHaveBeenCalled()
    },
    20_000,
  )
})
