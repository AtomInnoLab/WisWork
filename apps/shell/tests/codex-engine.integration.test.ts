import { isAbsolute } from 'node:path'
import { realpathSync } from 'node:fs'
import type { MessagesRequest } from '@wiswork/codex-bridge'
import { describe, expect, it, vi } from 'vitest'
import { suspendToolExecution } from '@wiswork/agent-core'
import type { RenderSlide, ShapeRenderNode } from '@wiswork/pptx-render'
import { createSlidesSkill } from '../../slides/src/renderer/ai/slides-skill'
import { executePreparedGeometryFamilyTransaction } from '../../slides/src/renderer/ai/presentation-geometry-transactions'
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

describe('real 0.147 production engine bridge', () => {
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
      const scripts = [
        `addText('title', '新人入职培训', {x: 90, y: 80, w: 1100, h: 100}); addText('body', '欢迎加入 WisWork', {x: 120, y: 250, w: 1040, h: 100}); return 'cover';`,
        `addText('title', '第一天安排', {x: 90, y: 70, w: 1100, h: 90}); addText('body', '认识团队\\n配置环境\\n了解工作方式', {x: 120, y: 200, w: 1040, h: 260}); return 'agenda';`,
        `addText('title', '开始协作', {x: 90, y: 70, w: 1100, h: 90}); addText('body', '主动沟通\\n记录决策\\n及时反馈', {x: 120, y: 200, w: 1040, h: 260}); return 'collaboration';`,
      ]
      const calls = [
        {
          carrier: 'read',
          toolName: 'plan_deck',
          input: {
            core_hook: '从第一天到独立协作',
            style: '简洁、清晰、友好',
            pages: [
              { title: '新人入职培训', brief: '欢迎', layout: 'cover' },
              { title: '第一天安排', brief: '安排', layout: 'content' },
              { title: '开始协作', brief: '协作', layout: 'content' },
            ],
          },
        },
        { carrier: 'propose', toolName: 'add_slide', input: { sourceIndex: 0 } },
        { carrier: 'propose', toolName: 'add_slide', input: { sourceIndex: 1 } },
        ...scripts.map((code, slideIndex) => ({
          carrier: 'propose',
          toolName: 'execute_slide_script',
          input: { slideIndex, code, explanation: `生成第 ${slideIndex + 1} 页` },
        })),
      ] as const
      const upstream = vi.fn(async (request: MessagesRequest) => {
        const next = calls[providerCalls++]
        if (!next) return finalResponse()
        const capability = request.system?.match(/pass capability ([A-Za-z0-9_-]{43})/)?.[1]
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
                            fontSizePx: 24,
                            color: '#111111',
                            bold: false,
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
              readOnlyHint: tool.name === 'plan_deck',
              destructiveHint: tool.name !== 'plan_deck',
            },
          })),
        callTool: vi.fn((_: unknown, call: any) => {
          if (call.name === 'plan_deck') return skill.executeTool(call)
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
        expect(slides).toHaveLength(3)
        expect(slideText).toEqual([
          expect.stringContaining('新人入职培训'),
          expect.stringContaining('第一天安排'),
          expect.stringContaining('开始协作'),
        ])
        expect(confirmations).toHaveLength(5)
        expect(transactionIds).toHaveLength(3)
        expect(new Set(transactionIds).size).toBe(3)
        expect(events.at(-1)).toEqual({ type: 'terminal', status: 'completed' })
        expect(diagnostics).toContain('gateway_tool_call_completed')
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
