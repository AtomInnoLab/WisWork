import { isAbsolute } from 'node:path'
import { realpathSync } from 'node:fs'
import type { MessagesRequest } from '@wiswork/codex-bridge'
import { describe, expect, it, vi } from 'vitest'
import { suspendToolExecution } from '@wiswork/agent-core'
import type { RenderSlide, ShapeRenderNode } from '@wiswork/pptx-render'
import { createSlidesSkill } from '../../slides/src/renderer/ai/slides-skill'
import { executePreparedGeometryFamilyTransaction } from '../../slides/src/renderer/ai/presentation-geometry-transactions'
import { writePresentationE2eArtifact } from '../../slides/tests/presentation-e2e-artifact'
import {
  buildDocumentToolInstructions,
  createProductionCodexBootstrap,
  safeTurnFailure,
  startBestEffortCodexInterrupt,
} from '../src/main/codex-engine'

it('gives Codex an exact carrier envelope and the registered document tool schemas', () => {
  const instructions = buildDocumentToolInstructions({
    credentials: { sessionId: 'private-session', secret: 'private-secret' },
    listTools: () => [
      {
        name: 'ask_clarification',
        description: 'Ask structured questions.',
        inputSchema: { type: 'object', properties: { questions: { type: 'array' } } },
        annotations: { readOnlyHint: true, destructiveHint: false },
      },
      {
        name: 'plan_deck',
        description: 'Plan the deck.',
        inputSchema: {
          type: 'object',
          properties: { pages: { type: 'array' } },
          required: ['pages'],
        },
        annotations: { readOnlyHint: true, destructiveHint: false },
      },
    ],
  } as never)
  expect(instructions).toContain('"name":"plan_deck"')
  expect(instructions).toContain('"carrier":"wiswork_read"')
  expect(instructions).toContain('arguments only inside input')
  expect(instructions).toContain(
    'When the host workflow makes you decide that user feedback is required',
  )
  expect(instructions).toContain('MUST call ask_clarification')
  expect(instructions).not.toContain('missing information would materially improve')
  expect(instructions).toContain('Never present those questions only as assistant prose')
  expect(instructions).toContain('continue the same task after its tool result')
  expect(instructions).not.toContain('private-session')
  expect(instructions).not.toContain('private-secret')
})

const configuredExecutable = process.env.WISWORK_CODEX_INTEGRATION_EXECUTABLE
const executable = configuredExecutable ? realpathSync(configuredExecutable) : undefined
const realIt = executable && isAbsolute(executable) ? it : it.skip
const realWisUsageToken = process.env.WISWORK_REAL_WISUSAGE_TOKEN
const realWisIt = executable && isAbsolute(executable) && realWisUsageToken ? it : it.skip

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

function toolResponse(
  code: string,
  reasoning: 'encrypted' | 'plaintext' = 'encrypted',
  toolUseId = 'custom_7',
): Response {
  return new Response(
    [
      'data: {"type":"message_start","message":{"id":"msg_tool","model":"openai/gpt-5.6-sol","usage":{"input_tokens":1}}}\n\n',
      ...(reasoning === 'encrypted'
        ? [
            'data: {"type":"content_block_start","index":0,"content_block":{"type":"redacted_thinking","data":"opaque-production-reasoning"}}\n\n',
          ]
        : [
            'data: {"type":"content_block_start","index":0,"content_block":{"type":"thinking","thinking":"private production prefix","signature":null}}\n\n',
            'data: {"type":"content_block_delta","index":0,"delta":{"type":"thinking_delta","thinking":" and suffix"}}\n\n',
            'data: {"type":"content_block_delta","index":0,"delta":{"type":"signature_delta","signature":"opaque-signature"}}\n\n',
          ]),
      'data: {"type":"content_block_stop","index":0}\n\n',
      `data: ${JSON.stringify({ type: 'content_block_start', index: 1, content_block: { type: 'tool_use', id: toolUseId, name: 'exec', input: {} } })}\n\n`,
      `data: ${JSON.stringify({ type: 'content_block_delta', index: 1, delta: { type: 'input_json_delta', partial_json: JSON.stringify({ code }) } })}\n\n`,
      'data: {"type":"content_block_stop","index":1}\n\n',
      'data: {"type":"message_delta","delta":{"stop_reason":"tool_use"},"usage":{"output_tokens":4,"output_tokens_details":{"reasoning_tokens":3}}}\n\n',
      'data: {"type":"message_stop"}\n\n',
    ].join(''),
    { status: 200, headers: { 'content-type': 'text/event-stream' } },
  )
}

function turnCapability(request: MessagesRequest): string | undefined {
  const matches = [
    ...JSON.stringify(request.messages).matchAll(
      /<wiswork_turn_capability>([A-Za-z0-9_-]{43})<\/wiswork_turn_capability>/g,
    ),
  ]
  return matches.at(-1)?.[1]
}

describe('real 0.147 production engine bridge', () => {
  realIt.each([
    { scenario: 'recovers after receiving rejected exec feedback', keepsRejecting: false },
    { scenario: 'stops after three rejected exec feedback results', keepsRejecting: true },
  ])(
    '$scenario through the real Office model/tool loop',
    async ({ keepsRejecting }) => {
      const diagnostics: string[] = []
      const events: unknown[] = []
      const crashed = vi.fn()
      const callTool = vi.fn(async () => ({
        output: '{"slideCount":1,"selectedSlideIndex":0}',
        summary: 'Read presentation state',
      }))
      let providerCalls = 0
      const receivedFeedback: unknown[] = []
      const upstream = vi.fn(async (request: MessagesRequest) => {
        providerCalls += 1
        if (providerCalls > 1) {
          // Inspect actual runtime output, not the host-authored exec input echoed in history.
          const feedback = request.messages
            .flatMap((message) => message.content)
            .find(
              (block) =>
                block.type === 'tool_result' &&
                block.tool_use_id === `office_exec_${providerCalls - 1}`,
            )
          expect(feedback).toBeDefined()
          receivedFeedback.push(feedback)
          if (keepsRejecting || providerCalls === 2) {
            expect(JSON.stringify(feedback?.content)).toContain('invalid_tool_input')
            expect(JSON.stringify(feedback?.content)).toContain('No document tool was executed')
            expect(callTool).not.toHaveBeenCalled()
          } else {
            expect(JSON.stringify(feedback?.content)).toContain('slideCount')
            expect(callTool).toHaveBeenCalledOnce()
            return finalResponse()
          }
        }
        const capability = turnCapability(request)
        expect(capability).toBeTruthy()
        const argumentsText = JSON.stringify({
          capability,
          callId: `office-read-${providerCalls}`,
          toolName: 'get_presentation_state',
          input: {},
        })
        const code =
          keepsRejecting || providerCalls === 1
            ? `const state = await tools.mcp__wiswork__wiswork_read(${argumentsText}); text(state);`
            : `text(await tools.mcp__wiswork__wiswork_read(${argumentsText}));`
        return toolResponse(code, 'encrypted', `office_exec_${providerCalls}`)
      })
      const engine = await createProductionCodexBootstrap({
        fetchWithAuth: upstream,
        diagnostics: (code) => diagnostics.push(code),
      }).start({ executablePath: executable!, onCrash: crashed })
      const session = {
        identity: {
          ownerId: 'office-owner',
          host: 'office-powerpoint',
          documentId: 'office-exec-recovery',
          sessionId: 'office-session',
          generation: 1,
        },
        credentials: { sessionId: 'office-session', secret: 'secret' },
        listTools: () => [
          {
            name: 'get_presentation_state',
            description: 'Read presentation state.',
            inputSchema: { type: 'object', additionalProperties: false },
            annotations: { readOnlyHint: true, destructiveHint: false },
          },
        ],
        callTool,
        cancelAll: vi.fn(() => 0),
        close: vi.fn(),
      } as any
      engine.registerDocument!({
        ownerId: 'office-owner',
        documentId: 'office-exec-recovery',
        host: 'office-powerpoint',
        generation: 1,
        session,
        onEvent: (event) => events.push(event),
      })
      try {
        const running = engine.startTurn({
          documentId: 'office-exec-recovery',
          host: 'office-powerpoint',
          generation: 1,
          text: '继续，先读取当前 PowerPoint 状态。',
        })
        if (keepsRejecting) {
          await expect(running).rejects.toThrow('enhanced_response_incompatible')
          expect(upstream).toHaveBeenCalledTimes(4)
          expect(receivedFeedback).toHaveLength(3)
          expect(callTool).not.toHaveBeenCalled()
          expect(diagnostics).toContain('responses_stream_tool_input_retry_limit_exceeded')
          expect(events).toContainEqual(
            expect.objectContaining({ type: 'terminal', status: 'failed' }),
          )
        } else {
          await running
          expect(upstream).toHaveBeenCalledTimes(3)
          expect(receivedFeedback).toHaveLength(2)
          expect(callTool).toHaveBeenCalledExactlyOnceWith(
            session.credentials,
            expect.objectContaining({
              id: 'office-read-2',
              name: 'get_presentation_state',
              input: {},
            }),
          )
          expect(diagnostics.filter((code) => code.startsWith('responses_stream_'))).toEqual([])
          expect(events.at(-1)).toEqual({ type: 'terminal', status: 'completed' })
        }
        expect(crashed).not.toHaveBeenCalled()
      } finally {
        await engine.close()
      }
    },
    30_000,
  )

  realIt(
    'keeps the app-server alive across consecutive authority-bound turns',
    async () => {
      const upstream = vi.fn(async () => finalResponse())
      const crashed = vi.fn()
      const diagnostics: string[] = []
      const engine = await createProductionCodexBootstrap({
        fetchWithAuth: upstream,
        diagnostics: (code) => diagnostics.push(code),
      }).start({ executablePath: executable!, onCrash: crashed })
      const session = {
        identity: {
          ownerId: 'slides-owner',
          host: 'slides',
          documentId: 'consecutive-slides',
          sessionId: 'consecutive-session',
          generation: 1,
        },
        credentials: { sessionId: 'consecutive-session', secret: 'secret' },
        listTools: () => [
          {
            name: 'plan_deck',
            description: 'Plan a presentation before editing.',
            inputSchema: { type: 'object', additionalProperties: true },
            annotations: { readOnlyHint: true, destructiveHint: false },
          },
        ],
        callTool: vi.fn(),
        cancelAll: vi.fn(() => 0),
        close: vi.fn(),
      } as any
      engine.registerDocument!({
        ownerId: 'slides-owner',
        documentId: 'consecutive-slides',
        host: 'slides',
        generation: 1,
        session,
      })
      try {
        await engine
          .startTurn({
            documentId: 'consecutive-slides',
            host: 'slides',
            generation: 1,
            text: '做一份完整产品发布会演示。',
          })
          .catch((error) => {
            throw new Error(`first_turn_failed:${diagnostics.join(',')}`, { cause: error })
          })
        await engine
          .startTurn({
            documentId: 'consecutive-slides',
            host: 'slides',
            generation: 1,
            text: 'Previous user: 做一份完整产品发布会演示。 Previous assistant: 请提供产品名称和受众。 Latest user: 你帮我决定吧。',
          })
          .catch((error) => {
            throw new Error(`second_turn_failed:${diagnostics.join(',')}`, { cause: error })
          })
        expect(upstream).toHaveBeenCalledTimes(2)
        const firstRequest = upstream.mock.calls[0]![0]
        const secondRequest = upstream.mock.calls[1]![0]
        expect(JSON.stringify(secondRequest.messages)).toContain('做一份完整产品发布会演示')
        expect(JSON.stringify(secondRequest.messages)).toContain('你帮我决定吧')
        expect(turnCapability(firstRequest)).toBeTruthy()
        expect(turnCapability(secondRequest)).toBeTruthy()
        expect(turnCapability(secondRequest)).not.toBe(turnCapability(firstRequest))
        expect(crashed).not.toHaveBeenCalled()
      } finally {
        await engine.close()
      }
    },
    20_000,
  )

  realWisIt(
    'accepts the live WisUsage stream shape for a Slides planning turn',
    async () => {
      const diagnostics: string[] = []
      const upstream = async (request: MessagesRequest, signal?: AbortSignal) =>
        fetch('https://wisusage.atominnolab.com/v1/messages', {
          method: 'POST',
          headers: {
            authorization: `Bearer ${realWisUsageToken}`,
            'content-type': 'application/json',
            'x-req-location': 'sg',
          },
          body: JSON.stringify(request),
          signal,
        })
      const engine = await createProductionCodexBootstrap({
        fetchWithAuth: upstream,
        diagnostics: (code) => diagnostics.push(code),
      }).start({ executablePath: executable!, onCrash: vi.fn() })
      const session = {
        identity: {
          ownerId: 'slides-owner',
          host: 'slides',
          documentId: 'live-slides-stream',
          sessionId: 'live-slides-session',
          generation: 1,
        },
        credentials: { sessionId: 'live-slides-session', secret: 'secret' },
        listTools: () => [
          {
            name: 'plan_deck',
            description: 'Plan a presentation before editing.',
            inputSchema: { type: 'object', additionalProperties: true },
            annotations: { readOnlyHint: true, destructiveHint: false },
          },
        ],
        callTool: vi.fn(async () => ({ output: '{"planned":true}', summary: 'planned' })),
        cancelAll: vi.fn(() => 0),
        close: vi.fn(),
      } as any
      engine.registerDocument!({
        ownerId: 'slides-owner',
        documentId: 'live-slides-stream',
        host: 'slides',
        generation: 1,
        session,
      })
      try {
        await engine
          .startTurn({
            documentId: 'live-slides-stream',
            host: 'slides',
            generation: 1,
            text: '做一份新人入职培训课件，先规划结构。',
          })
          .catch((error) => {
            throw new Error(`live_engine_failed:${diagnostics.join(',')}`, { cause: error })
          })
        const gatewayDiagnostics = diagnostics.filter(
          (code) =>
            code.startsWith('gateway_input_shape_') ||
            code.startsWith('gateway_input_aliases_') ||
            code.startsWith('gateway_tool_call_denied_'),
        )
        expect(session.callTool, gatewayDiagnostics.join(',')).toHaveBeenCalled()
        expect(diagnostics).not.toContain('gateway_tool_call_denied')
      } finally {
        await engine.close()
      }
    },
    65_000,
  )

  realIt(
    'fails the first deterministic provider protocol error without waiting for retry timeout',
    async () => {
      const diagnostics: string[] = []
      const upstream = vi.fn(
        async () =>
          new Response('data: {"type":"unknown_private_shape"}\n\n', {
            headers: { 'content-type': 'text/event-stream' },
          }),
      )
      const engine = await createProductionCodexBootstrap({
        fetchWithAuth: upstream,
        diagnostics: (code) => diagnostics.push(code),
      }).start({ executablePath: executable!, onCrash: vi.fn() })
      const session = {
        identity: {
          ownerId: 'owner',
          host: 'docs',
          documentId: 'invalid-stream',
          sessionId: 'session',
          generation: 1,
        },
        credentials: { sessionId: 'session', secret: 'secret' },
        listTools: () => [],
        callTool: vi.fn(),
        cancelAll: vi.fn(() => 0),
        close: vi.fn(),
      } as any
      engine.registerDocument!({
        ownerId: 'owner',
        documentId: 'invalid-stream',
        host: 'docs',
        generation: 1,
        session,
      })
      try {
        const started = Date.now()
        await expect(
          engine.startTurn({
            documentId: 'invalid-stream',
            host: 'docs',
            generation: 1,
            text: 'Reply OK.',
          }),
        ).rejects.toThrow('enhanced_response_incompatible')
        expect(Date.now() - started).toBeLessThan(5_000)
        expect(diagnostics).toContain('responses_stream_unsupported_messages_event')
      } finally {
        await engine.close()
      }
    },
    15_000,
  )

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
        const capability = turnCapability(request)
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
        const capability = turnCapability(request)
        expect(capability).toBeTruthy()
        return toolResponse(
          `text(await tools.mcp__wiswork__wiswork_read(${JSON.stringify({ capability, callId: 'read-1', toolName: 'read_blocks', input: {} })}))`,
          'plaintext',
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
    'keeps buffered tool input alive past both native idle deadlines in Slides',
    async () => {
      const diagnostics: string[] = []
      const events: unknown[] = []
      const crashed = vi.fn()
      let providerCalls = 0
      let argumentChunks = 0
      let inputCompleted = false
      let streamStartedAt = 0
      const callTool = vi.fn(async () => {
        expect(inputCompleted).toBe(true)
        expect(Date.now() - streamStartedAt).toBeGreaterThan(60_000)
        return { output: '{"slideCount":1}', summary: 'Read presentation' }
      })
      const upstream = vi.fn(async (request: MessagesRequest) => {
        providerCalls += 1
        if (providerCalls > 1) {
          expect(providerCalls).toBe(2)
          expect(callTool).toHaveBeenCalledOnce()
          expect(
            JSON.stringify(
              request.messages
                .flatMap((message) => message.content)
                .find((block) => block.type === 'tool_result' && block.tool_use_id === 'custom_7'),
            ),
          ).toContain('slideCount')
          return finalResponse()
        }
        const capability = turnCapability(request)
        expect(capability).toBeTruthy()
        const code = `text(await tools.mcp__wiswork__wiswork_read(${JSON.stringify({ capability, callId: 'buffered-read', toolName: 'read_presentation', input: {} })}))`
        const frames = (await toolResponse(code).text())
          .trim()
          .split('\n\n')
          .flatMap((frame) => {
            const data = JSON.parse(frame.slice('data: '.length))
            if (data.delta?.type !== 'input_json_delta') return [{ frame, delayed: false }]
            const input = data.delta.partial_json as string
            return Array.from({ length: 20 }, (_, index) => ({
              frame: `data: ${JSON.stringify({
                ...data,
                delta: {
                  ...data.delta,
                  partial_json: input.slice(
                    Math.floor((index * input.length) / 20),
                    Math.floor(((index + 1) * input.length) / 20),
                  ),
                },
              })}`,
              delayed: true,
            }))
          })
        let index = 0
        let cancelled = false
        let timer: ReturnType<typeof setTimeout> | undefined
        let resume: (() => void) | undefined
        streamStartedAt = Date.now()
        return new Response(
          new ReadableStream<Uint8Array>({
            async pull(controller) {
              const next = frames[index++]
              if (!next) return controller.close()
              if (next.delayed) {
                await new Promise<void>((resolve) => {
                  resume = resolve
                  timer = setTimeout(resolve, 3_300)
                })
                timer = undefined
                resume = undefined
                if (cancelled) return
                expect(callTool).not.toHaveBeenCalled()
                expect(events).toEqual([])
                argumentChunks += 1
                inputCompleted = argumentChunks === 20
              }
              controller.enqueue(new TextEncoder().encode(`${next.frame}\n\n`))
            },
            cancel() {
              cancelled = true
              if (timer) clearTimeout(timer)
              resume?.()
            },
          }),
          { status: 200, headers: { 'content-type': 'text/event-stream' } },
        )
      })
      const engine = await createProductionCodexBootstrap({
        fetchWithAuth: upstream,
        diagnostics: (code) => diagnostics.push(code),
      }).start({ executablePath: executable!, onCrash: crashed })
      const session = {
        identity: {
          ownerId: 'owner',
          host: 'slides',
          documentId: 'buffered-slides',
          sessionId: 'session',
          generation: 1,
        },
        credentials: { sessionId: 'session', secret: 'secret' },
        listTools: () => [
          {
            name: 'read_presentation',
            inputSchema: { type: 'object', additionalProperties: false },
            annotations: { readOnlyHint: true, destructiveHint: false },
          },
        ],
        callTool,
        cancelAll: vi.fn(() => 0),
        close: vi.fn(),
      } as any
      engine.registerDocument!({
        ownerId: 'owner',
        documentId: 'buffered-slides',
        host: 'slides',
        generation: 1,
        session,
        onEvent: (event) => events.push(event),
      })
      try {
        await engine
          .startTurn({
            documentId: 'buffered-slides',
            host: 'slides',
            generation: 1,
            text: 'Read the presentation.',
          })
          .catch((error) => {
            throw new Error(`buffered_turn_failed:${diagnostics.join(',')}`, { cause: error })
          })
        expect(argumentChunks).toBe(20)
        expect(upstream).toHaveBeenCalledTimes(2)
        expect(callTool).toHaveBeenCalledExactlyOnceWith(
          session.credentials,
          expect.objectContaining({ id: 'buffered-read', name: 'read_presentation', input: {} }),
        )
        expect(crashed).not.toHaveBeenCalled()
        expect(events.at(-1)).toEqual({ type: 'terminal', status: 'completed' })
        expect(diagnostics).toContain('gateway_tool_call_completed')
        expect(diagnostics.filter((code) => code.startsWith('responses_stream_'))).toEqual([])
        expect(diagnostics).not.toContain('codex_error')
      } finally {
        await engine.close()
      }
    },
    95_000,
  )

  realIt(
    'generates and verifies a three-page onboarding deck through the real Codex and Slides tool loop',
    async () => {
      const blankSlide = (): RenderSlide => ({
        widthPx: 1280,
        heightPx: 720,
        scale: 1,
        background: { kind: 'solid', color: '#FFFFFF' },
        nodes: [],
      })
      let slides = [blankSlide()]
      let providerCalls = 0
      let transactionSequence = 0
      const transactionIds: string[] = []
      const confirmations: string[] = []
      const pendingConfirmations = new Map<string, () => void>()
      const diagnostics: string[] = []
      const onboardingContract = {
        schemaVersion: 1,
        revision: 1,
        status: 'ready',
        prototypePages: [1, 2, 3],
        brief: {
          topic: '新人入职培训',
          audience: '新员工',
          occasion: '入职培训',
          desiredOutcome: '从第一天开始独立协作',
          language: '中文',
          pageCount: 3,
          aspectRatio: '16:9',
          sourceConstraints: [],
        },
        narrative: {
          coreHook: '从第一天到独立协作',
          opening: '欢迎新成员',
          development: '介绍第一天安排',
          tension: '建立协作习惯',
          resolution: '形成三项行动',
          closingAction: '主动沟通、记录决策、及时反馈',
        },
        visualSystem: {
          style: '深色协作主题',
          colors: { primary: '#0B1020', accent: '#66E3FF' },
          typography: { title: '32pt', body: '20pt' },
          safeMargin: '64px',
          grid: '12 columns',
          imageTreatment: '右侧裁切，单页一个焦点',
          chartTreatment: '直接标注',
          antiPatterns: ['不缩小正文', '不使用占位图片'],
        },
        slides: [
          {
            number: 1,
            title: '新人入职培训',
            role: '建立主题和欢迎氛围',
            claim: '从第一天，到独立协作',
            content: ['从第一天，到独立协作'],
            evidence: [],
            visualRoute: '右侧团队协作照片',
            layoutFamily: 'cover',
            focalVisual: '团队协作照片',
            density: 'low',
            assetIds: ['team-photo'],
            acceptance: [{ id: 'A1.1', criterion: '图片不遮挡文字' }],
          },
          {
            number: 2,
            title: '第一天安排',
            role: '让新人理解首日节奏',
            claim: '认识团队、配置环境和了解工作方式',
            content: ['认识团队', '配置环境', '了解工作方式'],
            evidence: [],
            visualRoute: '三节点横向时间线',
            layoutFamily: 'timeline',
            focalVisual: '首日时间线',
            density: 'medium',
            assetIds: [],
            acceptance: [{ id: 'A2.1', criterion: '三个节点按顺序排列' }],
          },
          {
            number: 3,
            title: '开始协作',
            role: '把培训内容转为行动',
            claim: '形成三项协作行为',
            content: ['主动沟通', '记录决策', '及时反馈'],
            evidence: [],
            visualRoute: '三张等宽行动卡片',
            layoutFamily: 'cards',
            focalVisual: '行动卡片',
            density: 'medium',
            assetIds: [],
            acceptance: [{ id: 'A3.1', criterion: '三项行动同等突出' }],
          },
        ],
        assets: [
          {
            id: 'team-photo',
            slideNumbers: [1],
            type: 'image',
            role: 'substantive',
            intent: '表现团队协作',
            source: 'https://images.example/team',
            crop: '右侧裁切',
            placement: '页面右侧',
            status: 'ready',
            localReference: 'https://images.example/team.jpg',
          },
        ],
        deckAcceptance: [{ id: 'D1', criterion: '三页均通过截图验收' }],
      }
      const calls = [
        {
          carrier: 'read',
          toolName: 'image_search',
          input: { query: 'modern creative team collaboration', maxResults: 3 },
        },
        { carrier: 'read', toolName: 'plan_deck', input: { contract: onboardingContract } },
        {
          carrier: 'propose',
          toolName: 'build_deck',
          input: {
            theme: { mode: 'dark', primary: '#0B1020', accent: '#66E3FF' },
            phase: 'prototype',
            page_indexes: [0, 1, 2],
            pages: [
              {
                layout: 'cover',
                kicker: 'WELCOME · WISWORK',
                title: '新人入职培训',
                body: ['从第一天，到独立协作'],
                evidence: [],
                imageUrl: 'https://images.example/team.jpg',
                imageAlt: '团队协作场景',
              },
              {
                layout: 'timeline',
                title: '第一天安排',
                body: ['认识团队', '配置环境', '了解工作方式'],
                evidence: [],
              },
              {
                layout: 'cards',
                title: '开始协作',
                body: ['主动沟通', '记录决策', '及时反馈'],
                evidence: [],
              },
            ],
          },
        },
        { carrier: 'read', toolName: 'screenshot_slide', input: { slideIndex: 0 } },
        { carrier: 'read', toolName: 'screenshot_slide', input: { slideIndex: 1 } },
        { carrier: 'read', toolName: 'screenshot_slide', input: { slideIndex: 2 } },
      ] as const
      const upstream = vi.fn(async (request: MessagesRequest) => {
        const next = calls[providerCalls++]
        if (!next) return finalResponse()
        const capability = turnCapability(request)
        expect(capability).toBeTruthy()
        const method = next.carrier === 'read' ? 'wiswork_read' : 'wiswork_propose'
        return toolResponse(
          `text(await tools.mcp__wiswork__${method}(${JSON.stringify({ capability, callId: `deck-${providerCalls}`, toolName: next.toolName, input: next.input })}))`,
          'plaintext',
          `custom_deck_${providerCalls}`,
        )
      })
      ;(globalThis as any).window = {
        slidesApi: {
          imageSearch: vi.fn(async () => ({
            images: [{ imageUrl: 'https://images.example/team.jpg', title: 'Team' }],
            method: 'serper',
          })),
          insertImageUrl: vi.fn(async ({ slideIndex }: { slideIndex: number }) => ({
            sourceId: `image-${slideIndex}`,
            slide: slides[slideIndex],
          })),
          addSlide: vi.fn(async ({ sourceIndex }: { sourceIndex: number }) => {
            const next = slides.slice()
            const index = sourceIndex + 1
            next.splice(index, 0, blankSlide())
            return { slides: next, index }
          }),
        },
      }
      let activeSlideIndex = 0
      const hostApi = {
        preparePresentationTarget: vi.fn(async (request: { slideIndex: number }) => {
          activeSlideIndex = request.slideIndex
          return {
            status: 'prepared' as const,
            expectedDeckRevision: `sha256:${String(transactionIds.length).padStart(64, '0')}`,
            target: {
              slideId: `ppt/slides/slide${request.slideIndex + 1}.xml`,
              expectedFingerprint: `sha256:${String(request.slideIndex + 1).padStart(64, '0')}`,
            },
          }
        }),
        cancelPresentationTransaction: vi.fn(async () => true),
        executePresentationTransaction: vi.fn(async (transaction: any) => {
          const slide = slides[activeSlideIndex]!
          const created = new Map<string, string>()
          let nodes = slide.nodes.slice()
          for (const operation of transaction.operations as any[]) {
            if (operation.kind === 'add_text_box') {
              const sourceId = `generated-${++transactionSequence}`
              created.set(operation.clientId, sourceId)
              nodes.push({
                id: `render-${sourceId}`,
                sourceId,
                type: 'text',
                box: {
                  x: operation.geometry.x,
                  y: operation.geometry.y,
                  w: operation.geometry.width,
                  h: operation.geometry.height,
                  rotationDeg: operation.geometry.rotation,
                  flipH: false,
                  flipV: false,
                  centerX: operation.geometry.x + operation.geometry.width / 2,
                  centerY: operation.geometry.y + operation.geometry.height / 2,
                },
                fill: { kind: 'none' },
                text: { lines: [], insets: { l: 0, t: 0, r: 0, b: 0 }, anchor: 'top' },
              } as ShapeRenderNode)
            } else if (operation.kind === 'set_text') {
              const sourceId = operation.target.createdByClientId
                ? created.get(operation.target.createdByClientId)
                : operation.target.elementId
              nodes = nodes.map((node) =>
                node.sourceId !== sourceId
                  ? node
                  : ({
                      ...node,
                      text: {
                        ...(node as ShapeRenderNode).text!,
                        lines: operation.paragraphs.map((paragraph: any) => ({
                          runs: paragraph.runs.map((run: any) => ({
                            text: run.text,
                            x: 0,
                            baselineY: 24,
                            fontFamily: 'Arial',
                            fontSizePx: Number(run.fontSize ?? 18) / 0.75,
                            color: run.color ?? '#111111',
                            bold: Boolean(run.bold),
                            italic: false,
                            underline: false,
                            widthPx: String(run.text).length * 12,
                          })),
                          top: 0,
                          height: 28,
                        })),
                      },
                    } as ShapeRenderNode),
              )
            } else if (operation.kind === 'set_fill') {
              const sourceId = operation.target.createdByClientId
                ? created.get(operation.target.createdByClientId)
                : operation.target.elementId
              nodes = nodes.map((node) =>
                node.sourceId !== sourceId
                  ? node
                  : ({
                      ...node,
                      fill: {
                        kind: 'solid',
                        color: operation.fill.color,
                        transparency: operation.fill.transparency ?? 0,
                      },
                    } as ShapeRenderNode),
              )
            }
          }
          slides = slides.map((candidate, index) =>
            index === activeSlideIndex ? { ...candidate, nodes } : candidate,
          )
          transactionIds.push(transaction.transactionId)
          return {
            status: 'applied' as const,
            transactionId: transaction.transactionId,
            resultingDeckRevision: `sha256:${String(transactionIds.length).padStart(64, '0')}`,
            operationCount: transaction.operations.length,
          }
        }),
      }
      const executePresentationOperation = vi.fn((request: any, signal?: AbortSignal) =>
        executePreparedGeometryFamilyTransaction(hostApi, request, signal, async () => true),
      )
      const skill = createSlidesSkill({
        getSlides: () => slides,
        getCurrent: () => 0,
        getSelectedIds: () => [],
        applySlide: (index, slide) => {
          slides = slides.map((candidate, candidateIndex) =>
            candidateIndex === index ? slide : candidate,
          )
        },
        applyDeck: (next) => {
          slides = next
        },
        executePresentationOperation,
        fitWidthPx: 1280,
        captureSlideScreenshot: vi.fn(async () => ({ base64: 'aGVsbG8=', mime: 'image/png' })),
        reviewPresentationScreenshot: vi.fn(async () => true),
      })
      const relevantTools = new Map(
        skill.tools
          .filter((tool) => calls.some((call) => call.toolName === tool.name))
          .map((tool) => [tool.name, tool]),
      )
      const engine = await createProductionCodexBootstrap({
        fetchWithAuth: upstream,
        diagnostics: (code) => diagnostics.push(code),
      }).start({ executablePath: executable!, onCrash: vi.fn() })
      const registered: any = {
        identity: {
          ownerId: 'slides-owner',
          host: 'slides',
          documentId: 'slides-onboarding-deck',
          sessionId: 'slides-session',
          generation: 1,
        },
        credentials: { sessionId: 'slides-session', secret: 'secret' },
        listTools: () =>
          [...relevantTools.values()].map((tool) => ({
            ...tool,
            annotations: {
              readOnlyHint: tool.name === 'plan_deck' || tool.name === 'image_search',
              destructiveHint: tool.name !== 'plan_deck' && tool.name !== 'image_search',
            },
          })),
        callTool: vi.fn((_: unknown, call: any) => {
          if (call.name === 'plan_deck' || call.name === 'image_search')
            return skill.executeTool(call)
          let confirm!: () => void
          const confirmed = new Promise<void>((resolve) => {
            confirm = resolve
          })
          const result = confirmed.then(() => skill.executeTool(call))
          pendingConfirmations.set(call.id, confirm)
          return suspendToolExecution(result)
        }),
        cancelAll: vi.fn(() => 0),
        close: vi.fn(),
      }
      const events: any[] = []
      engine.registerDocument!({
        ownerId: 'slides-owner',
        documentId: 'slides-onboarding-deck',
        host: 'slides',
        generation: 1,
        session: registered,
        summarizeProposal: () => ({
          operation: 'restructure',
          target: 'slides',
          scope: 'whole-document',
        }),
        onEvent: (event) => {
          events.push(event)
          if (event.type !== 'proposal') return
          const confirm = pendingConfirmations.get(event.call.id)
          if (!confirm) throw new Error('missing_test_confirmation')
          pendingConfirmations.delete(event.call.id)
          confirmations.push(event.call.id)
          confirm()
        },
      })
      try {
        await engine.startTurn({
          documentId: 'slides-onboarding-deck',
          host: 'slides',
          generation: 1,
          text: '做一份新人入职培训课件',
        })
        const slideText = slides.map((slide) =>
          slide.nodes
            .flatMap((node) =>
              node.type === 'text'
                ? (node.text?.lines.flatMap((line) => line.runs.map((run) => run.text)) ?? [])
                : [],
            )
            .join('\n'),
        )
        expect(
          slides,
          JSON.stringify({ providerCalls, confirmations, events, diagnostics }),
        ).toHaveLength(3)
        expect(
          slideText,
          JSON.stringify({
            requests: executePresentationOperation.mock.calls,
            transactions: hostApi.executePresentationTransaction.mock.calls,
          }),
        ).toEqual([
          expect.stringContaining('新人入职培训'),
          expect.stringContaining('第一天安排'),
          expect.stringContaining('开始协作'),
        ])
        expect(confirmations).toHaveLength(1)
        expect(transactionIds).toHaveLength(3)
        expect(new Set(transactionIds).size).toBe(3)
        expect((globalThis as any).window.slidesApi.imageSearch).toHaveBeenCalledOnce()
        expect((globalThis as any).window.slidesApi.insertImageUrl).toHaveBeenCalledWith(
          expect.objectContaining({ slideIndex: 0, url: 'https://images.example/team.jpg' }),
        )
        expect(
          slides.flatMap((slide) => slide.nodes).some((node) => node.fill.kind === 'solid'),
        ).toBe(true)
        expect(events.at(-1)).toEqual({ type: 'terminal', status: 'completed' })
        expect(diagnostics).toContain('gateway_tool_call_completed')
        const artifact = await writePresentationE2eArtifact(
          slides,
          process.env.WISWORK_ENHANCED_PPT_E2E_OUTPUT ?? '/tmp/wiswork-enhanced-ppt-e2e.pptx',
        )
        expect(artifact.slideCount).toBe(3)
        expect(artifact.text).toMatch(/新人入职培训/)
      } finally {
        await engine.close()
        delete (globalThis as any).window
      }
    },
    65_000,
  )

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
        expect(upstream.mock.calls[0]![0].tools?.map((tool) => tool.name).sort()).toEqual([
          'exec',
          'wait',
        ])
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
