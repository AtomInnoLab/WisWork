import { isAbsolute } from 'node:path'
import { realpathSync } from 'node:fs'
import type { MessagesRequest } from '@wiswork/codex-bridge'
import { describe, expect, it, vi } from 'vitest'
import { suspendToolExecution } from '@wiswork/agent-core'
import {
  createProductionCodexBootstrap,
  safeTurnFailure,
  startBestEffortCodexInterrupt,
} from '../src/main/codex-engine'

const configuredExecutable = process.env.WISWORK_CODEX_INTEGRATION_EXECUTABLE
const executable = configuredExecutable ? realpathSync(configuredExecutable) : undefined
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

it('maps only bounded app-server error categories to actionable public failures', () => {
  expect(safeTurnFailure({ error: { codexErrorInfo: 'unauthorized' } })).toBe(
    'enhanced_auth_required',
  )
  expect(
    safeTurnFailure({
      error: { codexErrorInfo: { responseTooManyFailedAttempts: { httpStatusCode: 503 } } },
    }),
  ).toBe('enhanced_service_unavailable')
  expect(
    safeTurnFailure({
      error: { codexErrorInfo: { responseStreamDisconnected: { httpStatusCode: null } } },
    }),
  ).toBe('enhanced_connection_failed')
  expect(safeTurnFailure({ error: { message: 'private detail' } })).toBe('enhanced_turn_failed')
})

function finalResponse(): Response {
  return new Response(
    [
      'data: {"type":"message_start","message":{"id":"msg_1","model":"openai/gpt-5.6-sol","usage":{"input_tokens":1}}}\n\n',
      'data: {"type":"content_block_start","index":0,"content_block":{"type":"text","text":""}}\n\n',
      'data: {"type":"content_block_delta","index":0,"delta":{"type":"text_delta","text":"done"}}\n\n',
      'data: {"type":"content_block_stop","index":0}\n\n',
      'data: {"type":"message_delta","delta":{"stop_reason":"end_turn"}}\n\n',
      'data: {"type":"message_stop"}\n\n',
    ].join(''),
    { status: 200, headers: { 'content-type': 'text/event-stream' } },
  )
}

function toolResponse(code: string): Response {
  return new Response(
    [
      'data: {"type":"message_start","message":{"id":"msg_tool","model":"openai/gpt-5.6-sol","usage":{"input_tokens":1}}}\n\n',
      'data: {"type":"content_block_start","index":0,"content_block":{"type":"tool_use","id":"custom_7","name":"exec","input":{}}}\n\n',
      `data: ${JSON.stringify({ type: 'content_block_delta', index: 0, delta: { type: 'input_json_delta', partial_json: JSON.stringify({ code }) } })}\n\n`,
      'data: {"type":"content_block_stop","index":0}\n\n',
      'data: {"type":"message_delta","delta":{"stop_reason":"tool_use"}}\n\n',
      'data: {"type":"message_stop"}\n\n',
    ].join(''),
    { status: 200, headers: { 'content-type': 'text/event-stream' } },
  )
}

describe('real 0.147 production engine bridge', () => {
  realIt(
    'holds the real provider terminal until a pending host proposal settles',
    async () => {
      let providerCalls = 0
      let settle!: (value: any) => void
      const writer = new Promise<any>((resolve) => {
        settle = resolve
      })
      const upstream = vi.fn(async (request: MessagesRequest) => {
        providerCalls += 1
        if (providerCalls > 1) return finalResponse()
        const capability = request.system?.match(/pass capability ([A-Za-z0-9_-]{43})/)?.[1]
        return toolResponse(
          `text(await tools.mcp__wiswork__wiswork_propose(${JSON.stringify({ capability, callId: 'mutation-1', toolName: 'replace_blocks', input: {} })}))`,
        )
      })
      const engine = await createProductionCodexBootstrap({ fetchWithAuth: upstream }).start({
        executablePath: executable!,
        onCrash: vi.fn(),
      })
      const events: any[] = []
      const session = {
        identity: {
          ownerId: 'owner',
          host: 'docs',
          documentId: 'hold-doc',
          sessionId: 'session',
          generation: 1,
        },
        credentials: { sessionId: 'session', secret: 'secret' },
        listTools: () => [{ name: 'replace_blocks', annotations: { destructiveHint: true } }],
        callTool: () => suspendToolExecution(writer),
        cancelAll: vi.fn(() => 0),
        close: vi.fn(),
      } as any
      engine.registerDocument!({
        ownerId: 'owner',
        documentId: 'hold-doc',
        host: 'docs',
        generation: 1,
        session,
        summarizeProposal: () => ({ operation: 'replace', target: 'blocks', scope: 'bounded-set' }),
        onEvent: (event) => events.push(event),
      })
      try {
        let finished = false
        const running = engine
          .startTurn({ documentId: 'hold-doc', host: 'docs', generation: 1, text: 'replace' })
          .then(() => {
            finished = true
          })
        await vi.waitFor(
          () => expect(events.some((event) => event.type === 'proposal')).toBe(true),
          { timeout: 30_000 },
        )
        await vi.waitFor(() => expect(providerCalls).toBeGreaterThan(1), { timeout: 30_000 })
        expect(finished).toBe(false)
        expect(events.some((event) => event.type === 'terminal')).toBe(false)
        settle({ output: 'applied', summary: 'replace', mutated: true })
        await running
        expect(events.at(-1)).toEqual({ type: 'terminal', status: 'completed' })
      } finally {
        await engine.close()
      }
    },
    65_000,
  )

  realIt(
    'drives one bounded Docs read through the real model/tool loop',
    async () => {
      let providerCalls = 0
      const diagnostics: string[] = []
      const upstream = vi.fn(async (request: MessagesRequest) => {
        providerCalls += 1
        if (providerCalls > 1) return finalResponse()
        const match = request.system?.match(/pass capability ([A-Za-z0-9_-]{43})/)
        expect(match?.[1]).toBeTruthy()
        return toolResponse(
          `text(await tools.mcp__wiswork__wiswork_read(${JSON.stringify({ capability: match![1], callId: 'read-1', toolName: 'read_blocks', input: {} })}))`,
        )
      })
      const engine = await createProductionCodexBootstrap({
        fetchWithAuth: upstream,
        diagnostics: (code) => diagnostics.push(code),
      }).start({ executablePath: executable!, onCrash: vi.fn() })
      const session = {
        identity: {
          ownerId: 'owner',
          host: 'docs',
          documentId: 'docs-read',
          sessionId: 'session',
          generation: 1,
        },
        credentials: { sessionId: 'session', secret: 'secret' },
        listTools: () => [
          {
            name: 'read_blocks',
            annotations: { readOnlyHint: true, destructiveHint: false },
          },
        ],
        callTool: vi.fn(async () => ({ output: '{"paragraphs":1}', summary: 'read blocks' })),
        cancelAll: vi.fn(() => 0),
        close: vi.fn(),
      } as any
      const events: unknown[] = []
      engine.registerDocument!({
        ownerId: 'owner',
        documentId: 'docs-read',
        host: 'docs',
        generation: 1,
        session,
        onEvent: (event) => events.push(event),
      })
      try {
        await Promise.race([
          engine
            .startTurn({
              documentId: 'docs-read',
              host: 'docs',
              generation: 1,
              text: 'Read the document.',
            })
            .catch((error) => {
              throw new Error(`read_turn_failed:${diagnostics.join(',')}`, { cause: error })
            }),
          new Promise<never>((_, reject) =>
            setTimeout(
              () =>
                reject(
                  new Error(
                    `read_turn_timeout:${providerCalls}:${session.callTool.mock.calls.length}:${diagnostics.join(',')}`,
                  ),
                ),
              60_000,
            ).unref(),
          ),
        ])
        expect(session.callTool).toHaveBeenCalledWith(
          session.credentials,
          expect.objectContaining({ id: 'read-1', name: 'read_blocks', input: {} }),
        )
        expect(diagnostics).toContain('gateway_tool_call_received')
        expect(diagnostics).toContain('gateway_tool_call_completed')
        expect(events).toContainEqual(
          expect.objectContaining({ type: 'terminal', status: 'completed' }),
        )
      } finally {
        await engine.close()
      }
    },
    65_000,
  )

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
        listTools: () => [
          { name: 'read_blocks', annotations: { readOnlyHint: true, destructiveHint: false } },
        ],
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
        listTools: () => [
          { name: 'read_blocks', annotations: { readOnlyHint: true, destructiveHint: false } },
        ],
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
