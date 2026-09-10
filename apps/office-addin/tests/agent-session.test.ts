import type { AgentStreamCallbacks, AgentTransport, ToolExecution } from '@wiswork/agent-core'
import { describe, expect, it, vi } from 'vitest'
import {
  bindAuthLoss,
  createOfficeAgentSession,
  presentationClarificationText,
} from '../src/agent/use-office-agent.js'
import type { ProposalDecision, StructuredProposal } from '../src/agent/proposal-controller.js'
import { createStructuredProposalController } from '../src/agent/proposal-controller.js'
import type { OfficePowerPointVisualReviewer } from '../src/skills/powerpoint/powerpoint-verification.js'
import { createPcBridgeAgentTransport, type OfficeToolActivity } from '../src/agent/transport.js'
import { createOfficeDiagnostics } from '../src/diagnostics/office-diagnostics.js'

function transportHarness() {
  let callbacks: AgentStreamCallbacks | undefined
  const cancel = vi.fn()
  const stream = vi.fn((_request: unknown, next: AgentStreamCallbacks) => {
    callbacks = next
    return { cancel }
  })
  const transport: AgentTransport = {
    stream,
  }
  return { transport, cancel, stream, callbacks: () => callbacks! }
}

function proposalsHarness() {
  let pending:
    | { id: string; operation: 'replace'; before: string; value: string; fingerprint: string }
    | undefined
  const listeners = new Set<() => void>()
  let settleDecision: ((value: ProposalDecision) => void) | undefined
  let decision = Promise.resolve<ProposalDecision>({ status: 'cancelled' })
  const clear = (value: ProposalDecision) => {
    pending = undefined
    settleDecision?.(value)
    settleDecision = undefined
    listeners.forEach((listener) => listener())
  }
  const controller = {
    pending: () => pending,
    subscribe(listener: () => void) {
      listeners.add(listener)
      return () => {
        listeners.delete(listener)
      }
    },
    waitForDecision: () => decision,
    propose: vi.fn(),
    confirm: vi.fn(async () => {
      if (!pending) throw new Error('no_pending_proposal')
      clear({ status: 'confirmed' })
    }),
    reject: vi.fn(() => {
      clear({ status: 'rejected' })
    }),
    newTurn: vi.fn(() => {
      clear({ status: 'cancelled' })
    }),
    logout: vi.fn(() => {
      clear({ status: 'cancelled' })
    }),
    destroyDocumentContext: vi.fn(() => {
      clear({ status: 'cancelled' })
    }),
  }
  return {
    controller,
    setPending() {
      pending = {
        id: 'p1',
        operation: 'replace' as const,
        before: 'old',
        value: 'new',
        fingerprint: 'x',
      }
      decision = new Promise((resolve) => {
        settleDecision = resolve
      })
      listeners.forEach((listener) => listener())
    },
    clearPending() {
      pending = undefined
      listeners.forEach((listener) => listener())
    },
  }
}

describe('presentation clarification display', () => {
  it('never exposes the internal scope control code to people', () => {
    expect(presentationClarificationText('presentation_scope_required', () => '需要补充信息')).toBe(
      '需要补充信息',
    )
    expect(presentationClarificationText('Which slide?', () => '需要补充信息')).toBe('Which slide?')
  })
})

describe('Office agent session', () => {
  it.each([
    'review_required',
    'prototype_required',
    'production_incomplete',
    'verification_failed',
    'visual_review_failed',
    'invalid_status',
    'review_not_pending',
    'acceptance_mismatch',
    'screenshot_required',
    'unknown_private',
  ])('retains only the safe design-contract failure code: %s', async (suffix) => {
    let handler: ((call: any) => Promise<{ output: string; isError?: boolean }>) | undefined
    const code = `design_contract_${suffix}`
    const diagnostics = createOfficeDiagnostics({ host: 'powerpoint', build: 'test' })
    const session = createOfficeAgentSession({
      transport: transportHarness().transport,
      skill: {
        id: 'test',
        systemPrompt: '',
        tools: [{ name: 'plan_deck', description: 'plan', inputSchema: { type: 'object' } }],
        executeTool: vi.fn(async () => ({
          output:
            suffix === 'visual_review_failed'
              ? `${code}: private contract https://secret.example`
              : ['review_required', 'invalid_status', 'unknown_private'].includes(suffix)
                ? JSON.stringify({
                    error: code,
                    contract: 'private contract https://secret.example',
                  })
                : code,
          isError: true,
          summary: 'plan',
        })),
      },
      proposals: proposalsHarness().controller,
      diagnostics,
      remoteTools: {
        setToolHandler: (next) => {
          handler = next
        },
      },
    })
    await handler!({
      turnId: 'turn_12345678',
      callId: 'call_plan123',
      generation: 3,
      toolName: 'plan_deck',
      input: {},
      signal: new AbortController().signal,
    })
    expect(diagnostics.snapshot().events).toEqual([
      expect.objectContaining({
        tool: 'plan_deck',
        error_code: suffix === 'unknown_private' ? 'agent_run_failed' : code,
      }),
    ])
    expect(diagnostics.exportJson()).not.toContain('private contract')
    expect(diagnostics.exportJson()).not.toContain('secret.example')
    session.dispose()
  })

  it.each([
    ['image_fetch_unavailable', '图片暂时无法获取'],
    ['image_limit', '图片超过大小限制'],
    ['image_mime_unsupported', '图片格式不受支持'],
    ['invalid_image', '图片数据无效'],
  ])(
    'records a PC-only image failure with its visible tool and safe code: %s',
    async (code, message) => {
      const base = {
        type: 'wiswork_tool_lifecycle',
        generation: 3,
        call_id: 'call_image123',
        tool_name: 'insert_web_image',
        started_at: Date.now(),
      }
      const frames = [
        { ...base, state: 'running' },
        { ...base, state: 'error', summary: code },
        { type: 'message_delta', delta: { stop_reason: 'end_turn' } },
      ]
      const executeTool = vi.fn()
      const diagnostics = createOfficeDiagnostics({ host: 'powerpoint', build: 'test' })
      const session = createOfficeAgentSession({
        transport: createPcBridgeAgentTransport({
          snapshot: () => ({ enhanced: { session_generation: 3 } }),
          authenticatedFetch: vi.fn(
            async () =>
              new Response(frames.map((frame) => `data: ${JSON.stringify(frame)}\n\n`).join('')),
          ),
        } as any),
        skill: {
          id: 'test',
          systemPrompt: '',
          tools: [
            { name: 'insert_web_image', description: 'image', inputSchema: { type: 'object' } },
          ],
          executeTool,
        },
        proposals: proposalsHarness().controller,
        diagnostics,
      })
      session.send('Insert cover image')
      await vi.waitFor(() => expect(session.snapshot().busy).toBe(false))
      expect(session.snapshot().timeline.find((event) => event.kind === 'tool')).toMatchObject({
        name: 'insert_web_image',
        summary: '插入网络图片未完成',
        state: 'error',
        output: `${message}（${code}）`,
      })
      expect(diagnostics.snapshot().events).toEqual([
        expect.objectContaining({ tool: 'insert_web_image', phase: 'tool', error_code: code }),
      ])
      expect(executeTool).not.toHaveBeenCalled()
      session.dispose()
    },
  )

  it.each(['remote-first', 'observation-first'] as const)(
    'records a failed image call only once (%s)',
    async (order) => {
      let handler: ((call: any) => Promise<{ output: string; isError?: boolean }>) | undefined
      let observe: ((event: OfficeToolActivity) => void) | undefined
      const harness = transportHarness()
      const diagnostics = createOfficeDiagnostics({ host: 'powerpoint', build: 'test' })
      const session = createOfficeAgentSession({
        transport: {
          ...harness.transport,
          setToolActivityHandler: (next) => {
            observe = next
          },
        },
        skill: {
          id: 'test',
          systemPrompt: '',
          tools: [{ name: 'insert-image', description: 'image', inputSchema: { type: 'object' } }],
          executeTool: vi.fn(async () => ({
            output: 'invalid_image',
            isError: true,
            summary: 'image',
          })),
        },
        proposals: proposalsHarness().controller,
        diagnostics,
        remoteTools: {
          setToolHandler: (next) => {
            handler = next
          },
        },
      })
      session.send('Insert image')
      await Promise.resolve()
      const base = { callId: 'call_image123', toolName: 'insert-image', startedAt: Date.now() }
      observe!({ ...base, state: 'running' })
      const fail = () => observe!({ ...base, state: 'error', summary: 'invalid_image' })
      if (order === 'observation-first') fail()
      await handler!({
        ...base,
        turnId: 'turn_12345678',
        generation: 3,
        input: {},
        signal: new AbortController().signal,
      })
      if (order === 'remote-first') fail()
      fail()
      expect(diagnostics.snapshot().events).toEqual([
        expect.objectContaining({ tool: 'insert-image', error_code: 'invalid_image' }),
      ])
      expect(session.snapshot().timeline.find((event) => event.kind === 'tool')).toMatchObject({
        summary: '插入图片未完成',
        output: '图片数据无效（invalid_image）',
      })
      session.dispose()
    },
  )

  it('shows and resolves one model-authored questionnaire question at a time', async () => {
    let handler: ((call: any) => Promise<{ output: string; isError?: boolean }>) | undefined
    const session = createOfficeAgentSession({
      transport: transportHarness().transport,
      skill: {
        id: 'test',
        systemPrompt: 'test',
        tools: [{ name: 'ask_clarification', description: 'ask', inputSchema: { type: 'object' } }],
        executeTool: vi.fn(),
      },
      proposals: proposalsHarness().controller,
      remoteTools: {
        setToolHandler: (next) => {
          handler = next
        },
      },
    })
    const result = handler!({
      turnId: 'turn_12345678',
      callId: 'call_12345678',
      generation: 1,
      toolName: 'ask_clarification',
      input: {
        questions: [{ id: 'audience', label: '面向谁？', options: ['客户', '内部团队'] }],
      },
      signal: new AbortController().signal,
    })

    await vi.waitFor(() => expect(session.snapshot().questionnaire).toHaveLength(1))
    session.answerQuestionnaire?.('面向谁？: 客户')
    await expect(result).resolves.toMatchObject({ output: expect.stringContaining('客户') })
  })

  it('continues to plan when the model tries to finish after the questionnaire', async () => {
    const harness = transportHarness()
    const session = createOfficeAgentSession({
      transport: harness.transport,
      skill: {
        id: 'test',
        systemPrompt: 'test',
        tools: [
          { name: 'ask_clarification', description: 'ask', inputSchema: { type: 'object' } },
          { name: 'plan_deck', description: 'plan', inputSchema: { type: 'object' } },
        ],
        executeTool: vi.fn(async () => ({ output: 'planned', mutated: false, summary: 'Planned' })),
      },
      proposals: proposalsHarness().controller,
    })

    session.send('Create a deck')
    await Promise.resolve()
    harness.callbacks().onToolCall({
      id: 'questionnaire',
      name: 'ask_clarification',
      input: { questions: [{ id: 'audience', label: 'Audience?', options: ['A', 'B'] }] },
    })
    harness.callbacks().onDone()
    await vi.waitFor(() => expect(session.snapshot().questionnaire).toHaveLength(1))
    session.answerQuestionnaire?.('Audience: A')
    await vi.waitFor(() => expect(harness.stream).toHaveBeenCalledTimes(2))

    harness.callbacks().onDelta('I have the answers.')
    harness.callbacks().onDone()

    await vi.waitFor(() => expect(harness.stream).toHaveBeenCalledTimes(3))
    expect(harness.stream.mock.calls[2]?.[0]).toMatchObject({
      messages: expect.arrayContaining([
        expect.objectContaining({
          role: 'user',
          text: expect.stringContaining('Continue the WisWork Slides workflow with plan_deck'),
        }),
      ]),
    })
  })

  it('bounds a batched PowerPoint questionnaire to the first question', async () => {
    let handler: ((call: any) => Promise<{ output: string; isError?: boolean }>) | undefined
    const proposals = proposalsHarness()
    const session = createOfficeAgentSession({
      transport: transportHarness().transport,
      skill: {
        id: 'test',
        systemPrompt: 'test',
        tools: [{ name: 'ask_clarification', description: 'ask', inputSchema: { type: 'object' } }],
        executeTool: vi.fn(),
      },
      proposals: proposals.controller,
      remoteTools: {
        setToolHandler: (next) => {
          handler = next
        },
      },
    })
    const result = handler!({
      turnId: 'turn_12345678',
      callId: 'call_12345678',
      generation: 1,
      toolName: 'ask_clarification',
      input: {
        questions: [
          { id: 'audience', label: '面向谁？', options: ['客户', '内部团队'] },
          { id: 'style', label: '什么风格？', options: ['简洁', '杂志感'] },
        ],
      },
      signal: new AbortController().signal,
    })
    await vi.waitFor(() => expect(session.snapshot().questionnaire).toHaveLength(1))
    session.answerQuestionnaire?.('面向谁？: 客户')
    await expect(result).resolves.toMatchObject({ output: expect.stringContaining('客户') })
    expect(session.snapshot().questionnaire).toBeUndefined()
  })

  it('accepts model-authored structured questionnaire options from Enhanced mode', async () => {
    let handler: ((call: any) => Promise<{ output: string; isError?: boolean }>) | undefined
    const proposals = proposalsHarness()
    const session = createOfficeAgentSession({
      transport: transportHarness().transport,
      skill: {
        id: 'test',
        systemPrompt: 'test',
        tools: [{ name: 'ask_clarification', description: 'ask', inputSchema: { type: 'object' } }],
        executeTool: vi.fn(),
      },
      proposals: proposals.controller,
      remoteTools: {
        setToolHandler: (next) => {
          handler = next
        },
      },
    })
    const result = handler!({
      turnId: 'turn_12345678',
      callId: 'call_12345678',
      generation: 1,
      toolName: 'ask_clarification',
      input: {
        questions: [
          {
            id: 'audience',
            label: '面向谁？',
            options: [
              { label: '客户', description: '对外介绍' },
              { label: '内部团队', description: '内部培训' },
            ],
          },
        ],
      },
      signal: new AbortController().signal,
    })
    await vi.waitFor(() =>
      expect(session.snapshot().questionnaire?.[0]?.options).toEqual(['客户', '内部团队']),
    )
    session.answerQuestionnaire?.('面向谁？: 客户')
    await expect(result).resolves.toMatchObject({ output: expect.stringContaining('客户') })
  })

  it('returns the safe verification location and repair guidance to the paired PC model', async () => {
    let handler: ((call: any) => Promise<{ output: string; isError?: boolean }>) | undefined
    const proposals = createStructuredProposalController()
    const errorLocation = 'PowerPoint.operations.1.set_shape_text_style.fontFamily'
    createOfficeAgentSession({
      transport: transportHarness().transport,
      skill: {
        id: 'test',
        systemPrompt: '',
        tools: [
          { name: 'execute_office_js', description: 'write', inputSchema: { type: 'object' } },
        ],
        executeTool: async () => {
          const proposal = proposals.propose({
            operation: 'execute_office_js',
            title: 'Style',
            preview: {},
            impact: { host: 'powerpoint', targets: ['slide'], count: 1 },
            fingerprint: 'v1',
            validate: () => true,
            execute: () => undefined,
            verify: () => {
              throw Object.assign(new Error('office_verify_failed'), {
                debugInfo: { errorLocation },
              })
            },
          })
          return { output: JSON.stringify({ proposalId: proposal.id }), summary: 'prepared' }
        },
      },
      proposals,
      remoteTools: {
        setToolHandler: (next) => {
          handler = next
        },
      },
    })
    const result = handler!({
      turnId: 'turn_12345678',
      callId: 'call_12345678',
      generation: 1,
      toolName: 'execute_office_js',
      input: {},
      signal: new AbortController().signal,
    })
    await vi.waitFor(() => expect(proposals.pending()).toBeDefined())
    await expect(proposals.confirm(proposals.pending()!.id)).rejects.toThrow('office_verify_failed')
    const failure = await result
    expect(failure.isError).toBe(true)
    expect(JSON.parse(failure.output)).toMatchObject({
      error: 'office_verify_failed',
      errorLocation,
      instruction: expect.stringContaining('preserve the current family'),
    })
  })

  it('invalidates a suspended remote proposal when cancelled before confirmation', async () => {
    let handler: ((call: any) => Promise<{ output: string; isError?: boolean }>) | undefined
    const proposals = proposalsHarness()
    createOfficeAgentSession({
      transport: transportHarness().transport,
      skill: {
        id: 'test',
        systemPrompt: 'test',
        tools: [{ name: 'write_document', description: 'write', inputSchema: { type: 'object' } }],
        executeTool: vi.fn(async () => {
          proposals.setPending()
          return { output: '{"proposalId":"p1"}', summary: 'proposed', mutated: false }
        }),
      },
      proposals: proposals.controller,
      remoteTools: {
        setToolHandler: (next) => {
          handler = next
        },
      },
    })
    const abort = new AbortController()
    const result = handler!({
      turnId: 'turn_12345678',
      callId: 'call_12345678',
      generation: 1,
      toolName: 'write_document',
      input: {},
      signal: abort.signal,
    })
    await vi.waitFor(() => expect(proposals.controller.pending()).toBeDefined())
    abort.abort()
    await expect(result).resolves.toEqual({ output: 'tool_execution_failed', isError: true })
    expect(proposals.controller.pending()).toBeUndefined()
    await expect(proposals.controller.confirm()).rejects.toThrow('no_pending_proposal')
  })

  it('routes paired Enhanced calls through the same host skill and revokes the handler on dispose', async () => {
    let handler: ((call: any) => Promise<{ output: string; isError?: boolean }>) | undefined
    const setToolHandler = vi.fn((next) => {
      handler = next
    })
    const executeTool = vi.fn(async () => ({ output: '{"title":"Doc"}', summary: 'read' }))
    const session = createOfficeAgentSession({
      transport: transportHarness().transport,
      skill: {
        id: 'test',
        systemPrompt: 'test',
        tools: [{ name: 'read_document', description: 'read', inputSchema: { type: 'object' } }],
        executeTool,
      },
      proposals: proposalsHarness().controller,
      remoteTools: { setToolHandler },
    })
    expect(
      await handler!({
        turnId: 'turn_12345678',
        callId: 'call_12345678',
        generation: 1,
        toolName: 'read_document',
        input: {},
        signal: new AbortController().signal,
      }),
    ).toEqual({ output: '{"title":"Doc"}' })
    expect(session.snapshot().timeline).toEqual(
      expect.arrayContaining([
        expect.objectContaining({
          kind: 'tool',
          callId: 'call_12345678',
          state: 'complete',
        }),
      ]),
    )
    expect(executeTool).toHaveBeenCalledWith(
      expect.objectContaining({ name: 'read_document' }),
      expect.any(AbortSignal),
    )
    expect(
      await handler!({
        turnId: 'turn_12345678',
        callId: 'call_unknown12',
        generation: 1,
        toolName: 'shell',
        input: {},
        signal: new AbortController().signal,
      }),
    ).toEqual({ output: 'unknown_tool', isError: true })
    session.dispose()
    expect(setToolHandler).toHaveBeenLastCalledWith(undefined)
  })

  it('interleaves Enhanced progress text with remote tool execution', async () => {
    let handler: ((call: any) => Promise<{ output: string; isError?: boolean }>) | undefined
    const harness = transportHarness()
    const session = createOfficeAgentSession({
      transport: harness.transport,
      skill: {
        id: 'test',
        systemPrompt: 'test',
        tools: [{ name: 'read_document', description: 'read', inputSchema: { type: 'object' } }],
        executeTool: vi.fn(async () => ({ output: 'ok', summary: 'read' })),
      },
      proposals: proposalsHarness().controller,
      remoteTools: {
        setToolHandler: (next) => {
          handler = next
        },
      },
    })

    session.send('美化文稿')
    await Promise.resolve()
    harness.callbacks().onDelta('先检查文稿。')
    await handler!({
      turnId: 'turn_12345678',
      callId: 'call_state_12345678',
      generation: 1,
      toolName: 'read_document',
      input: {},
      signal: new AbortController().signal,
    })
    harness.callbacks().onDelta('再调整版式。')
    await handler!({
      turnId: 'turn_12345678',
      callId: 'call_shapes_12345678',
      generation: 1,
      toolName: 'read_document',
      input: {},
      signal: new AbortController().signal,
    })
    harness.callbacks().onDelta('最后复查。')

    expect(session.snapshot().timeline.map((event) => event.kind)).toEqual([
      'user',
      'assistant',
      'tool',
      'assistant',
      'tool',
      'assistant',
    ])
    expect(
      session
        .snapshot()
        .timeline.filter((event) => event.kind === 'assistant')
        .map((event) => ('text' in event ? event.text : '')),
    ).toEqual(['先检查文稿。', '再调整版式。', '最后复查。'])
  })

  it.each(['complete', 'error'])(
    'interleaves observed PC retrieval %s without local execution',
    async (state) => {
      const executeTool = vi.fn(async () => ({ output: 'must not execute', summary: 'search' }))
      const base = {
        type: 'wiswork_tool_activity',
        generation: 3,
        call_id: 'call_search123',
        tool_name: 'image_search',
        started_at: 1000,
        query: 'volcano',
      }
      const frames = [
        { type: 'content_block_delta', delta: { type: 'text_delta', text: '先搜索图片。' } },
        { ...base, state: 'running' },
        {
          ...base,
          state,
          summary: state === 'complete' ? '1 result' : 'Retrieval unavailable',
          ...(state === 'complete'
            ? {
                result_count: 1,
                display: {
                  kind: 'images',
                  items: [{ title: 'Volcano', url: 'https://example.com/volcano' }],
                },
              }
            : {}),
        },
        { type: 'content_block_delta', delta: { type: 'text_delta', text: '再准备规划。' } },
        { type: 'message_delta', delta: { stop_reason: 'end_turn' } },
      ]
      const transport = createPcBridgeAgentTransport({
        snapshot: () => ({ enhanced: { session_generation: 3 } }),
        authenticatedFetch: vi.fn(
          async () =>
            new Response(frames.map((frame) => `data: ${JSON.stringify(frame)}\n\n`).join('')),
        ),
      } as any)
      const session = createOfficeAgentSession({
        transport,
        skill: {
          id: 'test',
          systemPrompt: '',
          tools: [{ name: 'image_search', description: 'images', inputSchema: { type: 'object' } }],
          executeTool,
        },
        proposals: proposalsHarness().controller,
      })
      session.send('Create a deck')
      await vi.waitFor(() => expect(session.snapshot().busy).toBe(false))
      expect(session.snapshot().timeline.map((event) => event.kind)).toEqual([
        'user',
        'assistant',
        'tool',
        'assistant',
      ])
      expect(
        session
          .snapshot()
          .timeline.filter((event) => event.kind === 'assistant')
          .map((event) => ('text' in event ? event.text : '')),
      ).toEqual(['先搜索图片。', '再准备规划。'])
      expect(session.snapshot().timeline[2]).toMatchObject({ name: 'image_search', state })
      if (state === 'complete')
        expect(session.snapshot().timeline[2]).toMatchObject({
          display: {
            kind: 'images',
            items: [{ title: 'Volcano', url: 'https://example.com/volcano' }],
          },
        })
      expect(executeTool).not.toHaveBeenCalled()
      session.dispose()
    },
  )

  it.each(['observation-first', 'execution-first', 'execution-start-first'] as const)(
    'enriches one remote card with canonical failure (%s)',
    async (order) => {
      let handler: ((call: any) => Promise<{ output: string; isError?: boolean }>) | undefined
      let observe: ((event: OfficeToolActivity) => void) | undefined
      const harness = transportHarness()
      const executeTool = vi.fn(async () => ({ output: 'full execution receipt', summary: 'read' }))
      const session = createOfficeAgentSession({
        transport: {
          ...harness.transport,
          setToolActivityHandler: (next) => {
            observe = next
          },
        },
        skill: {
          id: 'test',
          systemPrompt: '',
          tools: [
            { name: 'get_document_text', description: 'read', inputSchema: { type: 'object' } },
          ],
          executeTool,
        },
        proposals: proposalsHarness().controller,
        remoteTools: {
          setToolHandler: (next) => {
            handler = next
          },
        },
      })
      session.send('Read')
      await Promise.resolve()
      const base = { callId: 'call_document123', toolName: 'get_document_text', startedAt: 1000 }
      if (order === 'observation-first')
        observe!({ ...base, state: 'running' } as OfficeToolActivity)
      const execution = handler!({
        turnId: 'turn_12345678',
        callId: base.callId,
        generation: 3,
        toolName: base.toolName,
        input: {},
        signal: new AbortController().signal,
      })
      if (order === 'execution-start-first')
        observe!({ ...base, state: 'running' } as OfficeToolActivity)
      await execution
      if (order !== 'execution-first')
        expect(session.snapshot().timeline.find((event) => event.kind === 'tool')).toMatchObject({
          state: 'running',
          output: 'full execution receipt',
        })
      else observe!({ ...base, state: 'running' } as OfficeToolActivity)
      observe!({ ...base, state: 'error', summary: 'Tool failed' } as OfficeToolActivity)
      const cards = session.snapshot().timeline.filter((event) => event.kind === 'tool')
      expect(cards).toHaveLength(1)
      expect(cards[0]).toMatchObject({ state: 'error', output: 'full execution receipt' })
      expect(executeTool).toHaveBeenCalledOnce()
      session.dispose()
    },
  )

  it.each(['canonical-error', 'new-task', 'disposed'] as const)(
    'does not publish a late remote receipt after %s',
    async (reason) => {
      let handler: ((call: any) => Promise<{ output: string; isError?: boolean }>) | undefined
      let observe: ((event: OfficeToolActivity) => void) | undefined
      let finish!: (value: ToolExecution) => void
      const harness = transportHarness()
      const executeTool = vi.fn(
        () =>
          new Promise<ToolExecution>((resolve) => {
            finish = resolve
          }),
      )
      const session = createOfficeAgentSession({
        transport: {
          ...harness.transport,
          setToolActivityHandler: (next) => {
            observe = next
          },
        },
        skill: {
          id: 'test',
          systemPrompt: '',
          tools: [
            { name: 'get_document_text', description: 'read', inputSchema: { type: 'object' } },
          ],
          executeTool,
        },
        proposals: proposalsHarness().controller,
        remoteTools: {
          setToolHandler: (next) => {
            handler = next
          },
        },
      })
      session.send('Old task')
      await Promise.resolve()
      const base = { callId: 'call_document123', toolName: 'get_document_text', startedAt: 1000 }
      observe!({ ...base, state: 'running' })
      const execution = handler!({
        turnId: 'turn_12345678',
        callId: base.callId,
        generation: 3,
        toolName: base.toolName,
        input: {},
        signal: new AbortController().signal,
      })
      if (reason === 'canonical-error')
        observe!({ ...base, state: 'error', summary: 'Interrupted' })
      else if (reason === 'new-task') {
        session.newTask()
        session.send('New task')
      } else session.dispose()
      const before = session.snapshot()
      finish({ output: 'late private receipt', summary: 'success' })
      await execution
      expect(session.snapshot().timeline).toEqual(before.timeline)
      expect(JSON.stringify(session.snapshot())).not.toContain('late private receipt')
      expect(executeTool).toHaveBeenCalledOnce()
      session.dispose()
    },
  )

  it('marks observed retrieval failed when its stream fails before a result', async () => {
    const start = {
      type: 'wiswork_tool_activity',
      generation: 3,
      call_id: 'call_search123',
      tool_name: 'image_search',
      started_at: 1000,
      state: 'running',
    }
    const transport = createPcBridgeAgentTransport({
      snapshot: () => ({ enhanced: { session_generation: 3 } }),
      authenticatedFetch: vi.fn(
        async () => new Response(`data: ${JSON.stringify(start)}\n\ndata: {"type":"error"}\n\n`),
      ),
    })
    const session = createOfficeAgentSession({
      transport,
      skill: {
        id: 'test',
        systemPrompt: '',
        tools: [{ name: 'image_search', description: 'images', inputSchema: { type: 'object' } }],
        executeTool: vi.fn(),
      },
      proposals: proposalsHarness().controller,
    })
    session.send('Create a deck')
    await vi.waitFor(() => expect(session.snapshot().status).toBe('error'))
    expect(session.snapshot().timeline.find((event) => event.kind === 'tool')).toMatchObject({
      state: 'error',
    })
    session.dispose()
  })

  it('records paired Enhanced semantic tool failures in Taskpane diagnostics', async () => {
    let handler: ((call: any) => Promise<{ output: string; isError?: boolean }>) | undefined
    const diagnostics = {
      startTrace: vi.fn(() => 'trace'),
      setTool: vi.fn(),
      record: vi.fn(),
      clear: vi.fn(),
    }
    createOfficeAgentSession({
      transport: transportHarness().transport,
      skill: {
        id: 'test',
        systemPrompt: 'test',
        tools: [
          { name: 'list_slide_shapes', description: 'read', inputSchema: { type: 'object' } },
        ],
        executeTool: vi.fn(async () => ({
          output: 'office_read_failed',
          isError: true,
          mutated: false,
          summary: 'failed',
        })),
      },
      proposals: proposalsHarness().controller,
      diagnostics,
      remoteTools: {
        setToolHandler: (next) => {
          handler = next
        },
      },
    })
    await handler!({
      turnId: 'turn_12345678',
      callId: 'call_12345678',
      generation: 1,
      toolName: 'list_slide_shapes',
      input: { slide_index: 0 },
      signal: new AbortController().signal,
    })
    expect(diagnostics.setTool).toHaveBeenCalledWith('list_slide_shapes')
    expect(diagnostics.record).toHaveBeenCalledWith(
      expect.objectContaining({ phase: 'tool', errorCode: 'office_read_failed' }),
    )
  })

  it('preserves bounded local diagnostics when Relay authentication is lost', () => {
    const diagnostics = {
      startTrace: vi.fn(() => 'trace'),
      setTool: vi.fn(),
      record: vi.fn(),
      clear: vi.fn(),
    }
    const session = createOfficeAgentSession({
      transport: transportHarness().transport,
      skill: { id: 'test', systemPrompt: 'test', tools: [], executeTool: vi.fn() },
      proposals: proposalsHarness().controller,
      diagnostics,
    })

    session.authenticationLost()

    expect(diagnostics.clear).not.toHaveBeenCalled()
    session.logout()
    expect(diagnostics.clear).toHaveBeenCalledOnce()
  })

  it('correlates a run and records a stable tool failure without retaining output', async () => {
    const harness = transportHarness()
    const diagnostics = {
      startTrace: vi.fn(() => 'trace'),
      setTool: vi.fn(),
      record: vi.fn(),
      clear: vi.fn(),
    }
    const session = createOfficeAgentSession({
      transport: harness.transport,
      skill: {
        id: 'test',
        systemPrompt: 'test',
        tools: [{ name: 'write_document', description: 'write', inputSchema: { type: 'object' } }],
        executeTool: vi.fn(async () => ({
          output: 'office_api_unsupported',
          isError: true,
          summary: 'failed',
        })),
      },
      proposals: proposalsHarness().controller,
      diagnostics,
    })
    session.send('write my secret article')
    await Promise.resolve()
    harness.callbacks().onToolCall({ id: 'call', name: 'write_document', input: {} })
    harness.callbacks().onDone()
    await vi.waitFor(() => expect(diagnostics.record).toHaveBeenCalled())
    expect(diagnostics.startTrace).toHaveBeenCalledOnce()
    expect(diagnostics.setTool).toHaveBeenCalledWith('write_document')
    expect(diagnostics.record).toHaveBeenCalledWith({
      phase: 'tool',
      errorCode: 'office_api_unsupported',
      durationMs: expect.any(Number),
    })
    expect(JSON.stringify(diagnostics.record.mock.calls)).not.toContain('write my secret article')
  })

  it('preserves a safe image-fetch failure code in diagnostics', async () => {
    const harness = transportHarness()
    const diagnostics = {
      startTrace: vi.fn(() => 'trace'),
      setTool: vi.fn(),
      record: vi.fn(),
      clear: vi.fn(),
    }
    const session = createOfficeAgentSession({
      transport: harness.transport,
      skill: {
        id: 'powerpoint',
        systemPrompt: 'test',
        tools: [
          { name: 'insert_web_image', description: 'image', inputSchema: { type: 'object' } },
        ],
        executeTool: vi.fn(async () => ({
          output: 'image_fetch_unavailable',
          isError: true,
          summary: 'failed',
        })),
      },
      proposals: proposalsHarness().controller,
      diagnostics,
    })
    session.send('insert an image')
    await Promise.resolve()
    harness.callbacks().onToolCall({ id: 'call', name: 'insert_web_image', input: {} })
    harness.callbacks().onDone()
    await vi.waitFor(() => expect(diagnostics.record).toHaveBeenCalled())
    expect(diagnostics.record).toHaveBeenCalledWith({
      phase: 'tool',
      errorCode: 'image_fetch_unavailable',
      durationMs: expect.any(Number),
    })
  })

  it('forwards an in-memory Office diagnostic cause without adding it to model output', async () => {
    const harness = transportHarness()
    const officeError = Object.assign(new Error('secret workbook value'), {
      name: 'RichApi.Error',
      code: 'InvalidArgument',
      debugInfo: { errorLocation: 'Worksheet.getRange' },
    })
    const diagnostics = {
      startTrace: vi.fn(() => 'trace'),
      setTool: vi.fn(),
      record: vi.fn(),
      clear: vi.fn(),
    }
    const session = createOfficeAgentSession({
      transport: harness.transport,
      skill: {
        id: 'excel',
        systemPrompt: 'test',
        tools: [{ name: 'get_cell_ranges', description: 'read', inputSchema: { type: 'object' } }],
        executeTool: vi.fn(async () => ({
          output: 'office_read_failed',
          isError: true,
          summary: 'failed',
          diagnosticError: officeError,
        })),
      },
      proposals: proposalsHarness().controller,
      diagnostics,
    })
    session.send('read cells')
    await Promise.resolve()
    harness.callbacks().onToolCall({ id: 'call', name: 'get_cell_ranges', input: {} })
    harness.callbacks().onDone()

    await vi.waitFor(() =>
      expect(diagnostics.record).toHaveBeenCalledWith({
        phase: 'tool',
        errorCode: 'office_read_failed',
        error: officeError,
        durationMs: expect.any(Number),
      }),
    )
    await vi.waitFor(() => expect(harness.stream).toHaveBeenCalledTimes(2))
    expect(JSON.stringify(harness.stream.mock.calls[1])).not.toContain('secret workbook value')
  })
  it('keeps a bounded immutable two-turn user and assistant presentation timeline', async () => {
    const harness = transportHarness()
    const proposals = proposalsHarness()
    const session = createOfficeAgentSession({
      transport: harness.transport,
      skill: {
        id: 'test',
        systemPrompt: 'test',
        tools: [
          {
            name: 'propose_replace_selection',
            description: 'prepare a proposal',
            inputSchema: { type: 'object' },
          },
        ],
        executeTool: vi.fn(async () => ({ output: 'prepared', summary: 'Prepared edit' })),
      },
      proposals: proposals.controller,
    })

    session.send('First question')
    await Promise.resolve()
    harness.callbacks().onDelta('First')
    harness.callbacks().onDelta(' answer')
    harness.callbacks().onDone()
    session.send('Second question')
    await Promise.resolve()
    harness.callbacks().onDelta('Second answer')
    harness.callbacks().onDone()

    const timeline = session.snapshot().timeline
    expect(timeline.map(({ kind }) => kind)).toEqual(['user', 'assistant', 'user', 'assistant'])
    expect(timeline.map((event) => ('text' in event ? event.text : ''))).toEqual([
      'First question',
      'First answer',
      'Second question',
      'Second answer',
    ])
    expect(Object.isFrozen(timeline)).toBe(true)
    expect(Object.isFrozen(timeline[0])).toBe(true)
  })

  it('replaces only the active assistant event while streaming', async () => {
    const harness = transportHarness()
    const session = createOfficeAgentSession({
      transport: harness.transport,
      skill: { id: 'test', systemPrompt: 'test', tools: [], executeTool: vi.fn() },
      proposals: proposalsHarness().controller,
    })

    session.send('Stream')
    await Promise.resolve()
    harness.callbacks().onDelta('A')
    const first = session.snapshot().timeline
    harness.callbacks().onDelta('B')
    const second = session.snapshot().timeline

    expect(first).toHaveLength(2)
    expect(second).toHaveLength(2)
    expect(second[0]).toBe(first[0])
    expect(second[1]).toMatchObject({ id: first[1]?.id, kind: 'assistant', text: 'AB' })
  })

  it('places completed tool work and its proposal inline before the following assistant turn', async () => {
    const harness = transportHarness()
    const proposals = proposalsHarness()
    const session = createOfficeAgentSession({
      transport: harness.transport,
      skill: {
        id: 'test',
        systemPrompt: 'test',
        tools: [
          {
            name: 'propose_replace_selection',
            description: 'prepare a proposal',
            inputSchema: { type: 'object' },
          },
        ],
        executeTool: vi.fn(async () => {
          proposals.setPending()
          return { output: 'prepared', summary: 'Prepared edit' }
        }),
      },
      proposals: proposals.controller,
    })

    session.send('Edit this')
    await Promise.resolve()
    harness.callbacks().onDelta('I will prepare it.')
    harness.callbacks().onToolCall({ id: 'tool-1', name: 'propose_replace_selection', input: {} })
    harness.callbacks().onDone()
    await vi.waitFor(() => expect(session.snapshot().proposal?.id).toBe('p1'))
    await session.confirm('p1')
    await vi.waitFor(() => expect(harness.stream).toHaveBeenCalledTimes(2))
    harness.callbacks().onDelta('Ready for review.')
    harness.callbacks().onDone()

    expect(session.snapshot().timeline.map(({ kind }) => kind)).toEqual([
      'user',
      'assistant',
      'tool',
      'proposal',
      'assistant',
    ])
    expect(session.snapshot().timeline[2]).toMatchObject({
      kind: 'tool',
      callId: 'tool-1',
      state: 'complete',
    })
    expect(session.snapshot().timeline[3]).toMatchObject({
      kind: 'proposal',
      proposal: { id: 'p1' },
      state: 'applied',
    })
  })

  it('does not fabricate a thinking message when the model emits only a tool call', async () => {
    const harness = transportHarness()
    const session = createOfficeAgentSession({
      transport: harness.transport,
      skill: {
        id: 'test',
        systemPrompt: 'test',
        tools: [{ name: 'web_search', description: 'search', inputSchema: { type: 'object' } }],
        executeTool: vi.fn(async () => ({ output: 'ok', summary: 'searched' })),
      },
      proposals: proposalsHarness().controller,
    })

    session.send('Research and build')
    await Promise.resolve()
    harness.callbacks().onToolCall({ id: 'search-1', name: 'web_search', input: {} })
    harness.callbacks().onDone()
    await vi.waitFor(() => expect(harness.stream).toHaveBeenCalledTimes(2))

    expect(session.snapshot().timeline.map((event) => event.kind)).toEqual(['user', 'tool'])
  })

  it('never exposes internal tool identifiers while a tool is running or fails', async () => {
    const harness = transportHarness()
    let failTool!: () => void
    const session = createOfficeAgentSession({
      transport: harness.transport,
      skill: {
        id: 'test',
        systemPrompt: 'test',
        tools: [{ name: 'bash', description: 'internal', inputSchema: { type: 'object' } }],
        executeTool: vi.fn(
          () =>
            new Promise<ToolExecution>((resolve) => {
              failTool = () => resolve({ output: 'sandbox_denied', isError: true, summary: 'bash' })
            }),
        ),
      },
      proposals: proposalsHarness().controller,
    })

    session.send('Open my attachment')
    await Promise.resolve()
    harness.callbacks().onToolCall({ id: 'tool-private', name: 'bash', input: {} })
    harness.callbacks().onDone()
    await Promise.resolve()
    expect(JSON.stringify(session.snapshot())).not.toContain('Running bash')
    expect(session.snapshot().activity).toBe('正在处理附件…')
    failTool()
    await Promise.resolve()
    await Promise.resolve()
    expect(JSON.stringify(session.snapshot())).not.toContain('"summary":"bash"')
  })

  it('retries the last bounded instruction after a stable run error', async () => {
    const harness = transportHarness()
    const session = createOfficeAgentSession({
      transport: harness.transport,
      skill: { id: 'test', systemPrompt: 'test', tools: [], executeTool: vi.fn() },
      proposals: proposalsHarness().controller,
    })

    session.send('Try this')
    await Promise.resolve()
    harness.callbacks().onError('provider_unavailable')
    session.retry()
    await Promise.resolve()

    expect(session.snapshot().status).toBe('working')
    expect(
      session
        .snapshot()
        .timeline.filter((event) => event.kind === 'user')
        .map((event) => (event.kind === 'user' ? event.text : '')),
    ).toEqual(['Try this', 'Try this'])
  })

  it.each([
    ['authenticationLost', 'resolve'],
    ['newTask', 'reject'],
    ['logout', 'resolve'],
  ] as const)(
    'does not repopulate presentation after %s races an in-flight confirmation that later %s',
    async (reset, outcome) => {
      const harness = transportHarness()
      const proposals = proposalsHarness()
      proposals.setPending()
      let settle!: () => void
      proposals.controller.confirm.mockImplementation(
        () =>
          new Promise<void>((resolve, reject) => {
            settle = () => (outcome === 'resolve' ? resolve() : reject(new Error('proposal_stale')))
          }),
      )
      const session = createOfficeAgentSession({
        transport: harness.transport,
        skill: { id: 'test', systemPrompt: 'test', tools: [], executeTool: vi.fn() },
        proposals: proposals.controller,
      })

      const confirmation = session.confirm('p1')
      expect(session.snapshot().applying).toBe(true)
      session[reset]()
      expect(session.snapshot()).toMatchObject({
        applying: false,
        status: 'idle',
        activity: '',
        timeline: [],
        error: undefined,
      })

      settle()
      await confirmation
      expect(session.snapshot()).toMatchObject({
        applying: false,
        status: 'idle',
        activity: '',
        timeline: [],
        error: undefined,
      })
    },
  )

  it('renders and returns a quarantined pending write without claiming it was applied', async () => {
    const harness = transportHarness()
    const proposals = createStructuredProposalController()
    const presentationText = vi.fn((key: string) =>
      key === 'write_pending_quarantined' ? 'localized pending write quarantine' : key,
    )
    const session = createOfficeAgentSession({
      transport: harness.transport,
      skill: {
        id: 'word',
        systemPrompt: 'test',
        tools: [{ name: 'raw_write', description: 'write', inputSchema: { type: 'object' } }],
        executeTool: vi.fn(() => {
          const proposal = proposals.propose({
            operation: 'raw_write',
            title: 'Confirm raw write',
            preview: {},
            impact: { host: 'word', targets: ['document'], count: 1 },
            fingerprint: 'v1',
            validate: async () => true,
            execute: async () => undefined,
            verify: async () => {
              throw new Error('office_write_pending')
            },
          })
          return {
            output: JSON.stringify({ proposalId: proposal.id }),
            mutated: false,
            summary: 'Awaiting confirmation',
          }
        }),
      },
      proposals,
      presentationText: presentationText as never,
    })

    session.send('write')
    await Promise.resolve()
    harness.callbacks().onToolCall({ id: 'raw-1', name: 'raw_write', input: {} })
    harness.callbacks().onDone()
    await vi.waitFor(() => expect(session.snapshot().proposal).toBeDefined())
    await session.confirm(session.snapshot().proposal!.id)
    await vi.waitFor(() => expect(harness.stream).toHaveBeenCalledTimes(2))

    expect(session.snapshot().timeline).toEqual(
      expect.arrayContaining([
        expect.objectContaining({
          kind: 'proposal',
          state: 'uncertain',
          error: 'localized pending write quarantine',
        }),
      ]),
    )
    const resumed = harness.stream.mock.calls[1]?.[0] as {
      messages: Array<{ results?: Array<{ output: string }> }>
    }
    expect(JSON.parse(resumed.messages.at(-1)!.results![0]!.output)).toMatchObject({
      status: 'write_pending',
      safeCode: 'office_write_pending',
      instruction: expect.not.stringMatching(/was applied/i),
    })
  })

  it('rejects an in-loop proposal immediately and resumes without executing the write', async () => {
    const harness = transportHarness()
    const proposals = createStructuredProposalController()
    const execute = vi.fn(async () => undefined)
    const executeTool = vi.fn(() => {
      const proposal = proposals.propose({
        operation: 'write_document',
        title: 'Write document',
        preview: {},
        impact: { host: 'word', targets: ['document'], count: 1 },
        fingerprint: 'v1',
        validate: async () => true,
        execute,
      })
      return {
        output: JSON.stringify({ proposalId: proposal.id }),
        mutated: false,
        summary: 'Awaiting confirmation',
      }
    })
    const session = createOfficeAgentSession({
      transport: harness.transport,
      skill: {
        id: 'word',
        systemPrompt: 'test',
        tools: [{ name: 'write_document', description: 'write', inputSchema: { type: 'object' } }],
        executeTool,
      },
      proposals,
    })

    session.send('write')
    await Promise.resolve()
    harness.callbacks().onToolCall({ id: 'write-1', name: 'write_document', input: {} })
    harness.callbacks().onToolCall({ id: 'write-2', name: 'write_document', input: {} })
    harness.callbacks().onDone()
    await vi.waitFor(() => expect(session.snapshot().proposal).toBeDefined())

    session.reject()
    await vi.waitFor(() => expect(harness.stream).toHaveBeenCalledTimes(2))

    expect(execute).not.toHaveBeenCalled()
    expect(executeTool).toHaveBeenCalledOnce()
    expect(session.snapshot().timeline).toEqual(
      expect.arrayContaining([expect.objectContaining({ kind: 'proposal', state: 'rejected' })]),
    )
  })

  it('does not create another Word write proposal after stale validation in the same run', async () => {
    const harness = transportHarness()
    const proposals = createStructuredProposalController()
    const executeTool = vi.fn(() => {
      const proposal = proposals.propose({
        operation: 'write_document',
        title: 'Write document',
        preview: {},
        impact: { host: 'word', targets: ['document'], count: 1 },
        fingerprint: 'v1',
        validate: async () => false,
        execute: vi.fn(),
      })
      return {
        output: JSON.stringify({ proposalId: proposal.id }),
        mutated: false,
        summary: 'Awaiting confirmation',
      }
    })
    const session = createOfficeAgentSession({
      transport: harness.transport,
      skill: {
        id: 'word',
        systemPrompt: 'test',
        tools: [{ name: 'write_document', description: 'write', inputSchema: { type: 'object' } }],
        executeTool,
      },
      proposals,
    })

    session.send('write')
    await Promise.resolve()
    harness.callbacks().onToolCall({ id: 'write-1', name: 'write_document', input: {} })
    harness.callbacks().onDone()
    await vi.waitFor(() => expect(session.snapshot().proposal).toBeDefined())
    await session.confirm(session.snapshot().proposal!.id)
    await vi.waitFor(() => expect(harness.stream).toHaveBeenCalledTimes(2))

    harness.callbacks().onToolCall({ id: 'write-2', name: 'write_document', input: {} })
    harness.callbacks().onDone()
    await vi.waitFor(() => expect(harness.stream).toHaveBeenCalledTimes(3))
    expect(executeTool).toHaveBeenCalledOnce()
    expect(session.snapshot().proposal).toBeUndefined()

    harness.callbacks().onDelta('The document changed before the edit could be applied.')
    harness.callbacks().onDone()
    await vi.waitFor(() => expect(session.snapshot().busy).toBe(false))
    expect(executeTool).toHaveBeenCalledOnce()
    expect(session.snapshot().proposal).toBeUndefined()
  })

  it('maps arbitrary transport failures to a stable code, safe copy, and retry policy', async () => {
    const harness = transportHarness()
    const session = createOfficeAgentSession({
      transport: harness.transport,
      skill: { id: 'test', systemPrompt: 'test', tools: [], executeTool: vi.fn() },
      proposals: proposalsHarness().controller,
    })

    session.send('Fail safely')
    await Promise.resolve()
    harness.callbacks().onError('/Users/alice/private token=secret')

    expect(session.snapshot()).toMatchObject({
      error: 'agent_run_failed',
      errorMessage: 'The Agent could not complete this request. Try again.',
      retryable: true,
    })
    expect(JSON.stringify(session.snapshot())).not.toContain('alice')
    expect(JSON.stringify(session.snapshot())).not.toContain('secret')
  })

  it('reports the bounded transport deadline as a retryable request timeout', async () => {
    const harness = transportHarness()
    const session = createOfficeAgentSession({
      transport: harness.transport,
      skill: { id: 'test', systemPrompt: 'test', tools: [], executeTool: vi.fn() },
      proposals: proposalsHarness().controller,
    })

    session.send('Build a complex presentation')
    await Promise.resolve()
    harness.callbacks().onError('transport_timeout')

    expect(session.snapshot()).toMatchObject({
      error: 'request_timeout',
      errorMessage: 'The Agent took too long to respond. Try again.',
      retryable: true,
    })
  })

  it('does not retry a known non-retryable authentication failure', async () => {
    const harness = transportHarness()
    const session = createOfficeAgentSession({
      transport: harness.transport,
      skill: { id: 'test', systemPrompt: 'test', tools: [], executeTool: vi.fn() },
      proposals: proposalsHarness().controller,
    })
    session.send('Protected request')
    await Promise.resolve()
    harness.callbacks().onError('auth_required')
    session.retry()
    expect(session.snapshot()).toMatchObject({ error: 'auth_required', retryable: false })
    expect(session.snapshot().timeline.filter((event) => event.kind === 'user')).toHaveLength(1)
  })

  it('clears presentation state atomically for new task and logout', async () => {
    const harness = transportHarness()
    const proposals = proposalsHarness()
    const session = createOfficeAgentSession({
      transport: harness.transport,
      skill: { id: 'test', systemPrompt: 'test', tools: [], executeTool: vi.fn() },
      proposals: proposals.controller,
    })

    session.send('Old task')
    await Promise.resolve()
    harness.callbacks().onDelta('Old answer')
    harness.callbacks().onDone()
    session.newTask()
    expect(session.snapshot()).toMatchObject({ timeline: [], status: 'idle', error: undefined })
    expect(proposals.controller.logout).toHaveBeenCalledOnce()

    session.send('Another task')
    await Promise.resolve()
    session.logout()
    expect(harness.cancel).toHaveBeenCalledTimes(2)
    expect(session.snapshot().timeline).toEqual([])
  })

  it('preserves Agent history after stop but clears it for a new task', async () => {
    const harness = transportHarness()
    const session = createOfficeAgentSession({
      transport: harness.transport,
      skill: { id: 'test', systemPrompt: 'test', tools: [], executeTool: vi.fn() },
      proposals: proposalsHarness().controller,
    })

    session.send('Keep this context')
    await Promise.resolve()
    session.stop()
    harness.callbacks().onDone()
    await Promise.resolve()
    session.send('Continue')
    await Promise.resolve()

    expect(harness.stream.mock.calls[1]?.[0]).toMatchObject({
      messages: expect.arrayContaining([
        expect.objectContaining({ role: 'user', text: 'Keep this context' }),
        expect.objectContaining({ role: 'user', text: 'Continue' }),
      ]),
    })

    harness.callbacks().onDone()
    await Promise.resolve()
    session.newTask()
    session.send('Fresh context')
    await Promise.resolve()

    const freshRequest = harness.stream.mock.calls[2]?.[0] as {
      messages: Array<{ role: string; text?: string }>
    }
    expect(freshRequest.messages).toEqual([
      expect.objectContaining({ role: 'user', text: 'Fresh context' }),
    ])
  })

  it.each(['logout', 'dispose'] as const)(
    'suppresses late transport and proposal callbacks after %s',
    async (endSession) => {
      const harness = transportHarness()
      const proposals = proposalsHarness()
      const listener = vi.fn()
      const session = createOfficeAgentSession({
        transport: harness.transport,
        skill: { id: 'test', systemPrompt: 'test', tools: [], executeTool: vi.fn() },
        proposals: proposals.controller,
      })
      session.subscribe(listener)
      session.send('Old request')
      await Promise.resolve()
      const callbacks = harness.callbacks()

      session[endSession]()
      listener.mockClear()
      callbacks.onDelta('Late answer')
      callbacks.onDone()
      if (endSession === 'dispose') proposals.setPending()
      await Promise.resolve()

      expect(session.snapshot()).toMatchObject({
        assistantText: '',
        busy: false,
        timeline: [],
        proposal: undefined,
      })
      expect(listener).not.toHaveBeenCalled()
      if (endSession === 'dispose') {
        session.send('Must not run')
        session.stop()
        session.reject()
        session.newTask()
        session.retry()
        session.logout()
        session.authenticationLost()
        await Promise.resolve()
        expect(harness.stream).toHaveBeenCalledOnce()
        expect(proposals.controller.logout).toHaveBeenCalledOnce()
        expect(proposals.controller.reject).not.toHaveBeenCalled()
      }
    },
  )
  it('exposes generic structured proposal fields without legacy coercion', () => {
    const harness = transportHarness()
    const proposal: StructuredProposal = Object.freeze({
      id: 'structured',
      operation: 'edit_slide_xml',
      toolName: 'edit_slide_xml',
      title: 'Update slide XML',
      preview: { nodes: 2 },
      impact: { host: 'powerpoint', targets: ['slide-1'], count: 1 },
      fingerprint: 'fp',
      before: '<old/>',
      after: '<new/>',
      code: 'context.sync()',
    })
    const controller = {
      pending: () => proposal,
      subscribe: () => () => undefined,
      waitForDecision: vi.fn(async () => ({ status: 'cancelled' as const })),
      propose: vi.fn(),
      confirm: vi.fn(),
      reject: vi.fn(),
      newTurn: vi.fn(),
      logout: vi.fn(),
      destroyDocumentContext: vi.fn(),
      isQuarantined: vi.fn(() => false),
      quarantine: vi.fn(),
      resolveQuarantine: vi.fn(),
    }
    const session = createOfficeAgentSession({
      transport: harness.transport,
      skill: { id: 'test', systemPrompt: 'test', tools: [], executeTool: vi.fn() },
      proposals: controller,
    })
    expect(session.snapshot().proposal).toEqual(proposal)
  })

  it('streams assistant text and reports completion', async () => {
    const harness = transportHarness()
    const proposals = proposalsHarness()
    const session = createOfficeAgentSession({
      transport: harness.transport,
      skill: { id: 'test', systemPrompt: 'test', tools: [], executeTool: vi.fn() },
      proposals: proposals.controller,
    })

    session.send('Summarize this')
    await Promise.resolve()
    harness.callbacks().onDelta('Hello')
    harness.callbacks().onDelta(' world')
    harness.callbacks().onDone()

    expect(session.snapshot()).toMatchObject({
      assistantText: 'Hello world',
      busy: false,
      status: 'done',
    })
  })

  it('invalidates a pending proposal before a new instruction and on logout', async () => {
    const harness = transportHarness()
    const proposals = proposalsHarness()
    const session = createOfficeAgentSession({
      transport: harness.transport,
      skill: { id: 'test', systemPrompt: 'test', tools: [], executeTool: vi.fn() },
      proposals: proposals.controller,
    })

    proposals.setPending()
    session.send('new request')
    await Promise.resolve()
    expect(proposals.controller.newTurn).toHaveBeenCalledOnce()
    harness.callbacks().onDone()
    proposals.setPending()
    session.logout()

    expect(proposals.controller.logout).toHaveBeenCalledOnce()
    expect(session.snapshot()).toMatchObject({
      assistantText: '',
      proposal: undefined,
      busy: false,
    })
  })

  it('stops the active stream and surfaces safe proposal confirmation errors', async () => {
    const harness = transportHarness()
    const proposals = proposalsHarness()
    proposals.controller.confirm.mockImplementation(async () => {
      proposals.clearPending()
      throw new Error('proposal_stale')
    })
    proposals.setPending()
    const session = createOfficeAgentSession({
      transport: harness.transport,
      skill: { id: 'test', systemPrompt: 'test', tools: [], executeTool: vi.fn() },
      proposals: proposals.controller,
    })

    session.send('work')
    await Promise.resolve()
    session.stop()
    expect(harness.cancel).toHaveBeenCalledOnce()
    harness.callbacks().onDone()
    await session.confirm('p1')
    expect(session.snapshot()).toMatchObject({
      error: 'proposal_stale',
      errorMessage: '文档内容已发生变化，刚才的修改未应用。',
      retryable: true,
    })
    expect(session.snapshot().proposal).toBeUndefined()
  })

  it.each([
    ['office_verify_failed', 'The approved change could not be verified.'],
    [
      'office_overwrite_required',
      'The target cells contain data. Choose an empty range or explicitly allow overwrite.',
    ],
    ['office_recovery_failed', 'The document could not be restored after the failed change.'],
    [
      'office_concurrent_change',
      'The document changed during the operation. Inspect it before trying again.',
    ],
    [
      'office_state_uncertain',
      'The change may be partially applied. Wait for reconciliation; if editing stays blocked, reload the document before trying again.',
    ],
    [
      'office_recovery_failed:word_body_shape',
      'The document could not be restored after the failed change (word_body_shape).',
    ],
  ])('preserves the terminal confirmation code %s with safe copy', async (code, message) => {
    const proposals = proposalsHarness()
    proposals.controller.confirm.mockRejectedValue(new Error(code))
    proposals.setPending()
    const session = createOfficeAgentSession({
      transport: transportHarness().transport,
      skill: { id: 'test', systemPrompt: 'test', tools: [], executeTool: vi.fn() },
      proposals: proposals.controller,
    })

    await session.confirm('p1')

    expect(session.snapshot()).toMatchObject({
      error: code,
      errorMessage: message,
      retryable: false,
    })
  })

  it('allows only one confirmation, blocks competing writes, and lets Stop abort applying', async () => {
    const harness = transportHarness()
    const proposals = proposalsHarness()
    proposals.setPending()
    let settle!: () => void
    proposals.controller.confirm.mockImplementation(
      () =>
        new Promise<void>((resolve) => {
          settle = () => {
            proposals.clearPending()
            resolve()
          }
        }),
    )
    const session = createOfficeAgentSession({
      transport: harness.transport,
      skill: { id: 'test', systemPrompt: 'test', tools: [], executeTool: vi.fn() },
      proposals: proposals.controller,
    })

    const first = session.confirm('p1')
    const second = session.confirm('p1')
    session.send('race')
    session.reject()
    session.stop()

    expect(session.snapshot()).toMatchObject({ applying: false, status: 'cancelled' })
    expect(proposals.controller.confirm).toHaveBeenCalledOnce()
    expect(proposals.controller.newTurn).toHaveBeenCalledOnce()
    expect(proposals.controller.reject).not.toHaveBeenCalled()
    expect(harness.cancel).not.toHaveBeenCalled()

    settle()
    await Promise.all([first, second])
    expect(session.snapshot().applying).toBe(false)
    expect(session.snapshot().proposal).toBeUndefined()
    session.logout()
    expect(proposals.controller.logout).toHaveBeenCalledOnce()
  })

  it('pauses the same agent turn for approval, applies immediately, then resumes once', async () => {
    const harness = transportHarness()
    const proposals = proposalsHarness()
    const session = createOfficeAgentSession({
      transport: harness.transport,
      skill: {
        id: 'test',
        systemPrompt: 'test',
        tools: [{ name: 'write_document', description: 'write', inputSchema: { type: 'object' } }],
        executeTool: vi.fn(() => {
          proposals.setPending()
          return {
            output: JSON.stringify({ status: 'awaiting_user_confirmation' }),
            mutated: false,
            summary: 'Awaiting confirmation',
          }
        }),
      },
      proposals: proposals.controller,
    })

    session.send('prepare an edit')
    await Promise.resolve()
    harness.callbacks().onToolCall({ id: 'write-1', name: 'write_document', input: {} })
    harness.callbacks().onDone()
    await vi.waitFor(() => expect(session.snapshot().proposal?.id).toBe('p1'))

    expect(session.snapshot()).toMatchObject({ busy: true, applying: false })
    expect(harness.stream).toHaveBeenCalledTimes(1)
    await session.confirm('p1')
    expect(proposals.controller.confirm).toHaveBeenCalledOnce()
    await vi.waitFor(() => expect(harness.stream).toHaveBeenCalledTimes(2))
    expect(session.snapshot().timeline).toEqual(
      expect.arrayContaining([expect.objectContaining({ kind: 'proposal', state: 'applied' })]),
    )

    harness.callbacks().onDelta('The approved change is now applied.')
    harness.callbacks().onDone()
    await vi.waitFor(() => expect(session.snapshot().busy).toBe(false))
    expect(harness.stream).toHaveBeenCalledTimes(2)
  })

  it('auto-applies ordinary PowerPoint proposals without exposing confirmation UI', async () => {
    const harness = transportHarness()
    const proposals = createStructuredProposalController()
    const execute = vi.fn(async () => undefined)
    const session = createOfficeAgentSession({
      transport: harness.transport,
      skill: {
        id: 'powerpoint',
        systemPrompt: 'test',
        tools: [{ name: 'edit_slide_text', description: 'write', inputSchema: { type: 'object' } }],
        executeTool: vi.fn(() => {
          const proposal = proposals.propose({
            operation: 'edit_slide_text',
            toolName: 'edit_slide_text',
            title: 'Update slide',
            preview: {},
            impact: { host: 'powerpoint', targets: ['slide-1/shape-1'], count: 1 },
            fingerprint: 'v1',
            validate: async () => true,
            execute,
          })
          return {
            output: JSON.stringify({ proposalId: proposal.id }),
            mutated: false,
            summary: 'Prepared change',
          }
        }),
      },
      proposals,
      automaticPowerPointMutations: true,
    })

    session.send('update the slide')
    await Promise.resolve()
    harness.callbacks().onToolCall({ id: 'ppt-write', name: 'edit_slide_text', input: {} })
    harness.callbacks().onDone()

    await vi.waitFor(() => expect(execute).toHaveBeenCalledOnce())
    await vi.waitFor(() => expect(harness.stream).toHaveBeenCalledTimes(2))
    expect(session.snapshot().proposal).toBeUndefined()
    expect(session.snapshot().timeline.some((event) => event.kind === 'proposal')).toBe(false)
  })

  it('auto-applies a PowerPoint background proposal without blocking the Agent turn', async () => {
    const harness = transportHarness()
    const proposals = createStructuredProposalController()
    const execute = vi.fn(async () => undefined)
    const session = createOfficeAgentSession({
      transport: harness.transport,
      skill: {
        id: 'powerpoint',
        systemPrompt: 'test',
        tools: [
          { name: 'set_slide_background', description: 'write', inputSchema: { type: 'object' } },
        ],
        executeTool: vi.fn(() => {
          const proposal = proposals.propose({
            operation: 'set_slide_background',
            toolName: 'set_slide_background',
            title: 'Set slide background',
            preview: {},
            impact: { host: 'powerpoint', targets: ['slide-1/background'], count: 1 },
            fingerprint: 'v1',
            validate: async () => true,
            execute,
          })
          return {
            output: JSON.stringify({ proposalId: proposal.id }),
            mutated: false,
            summary: 'Prepared background change',
          }
        }),
      },
      proposals,
      automaticPowerPointMutations: true,
    })

    session.send('set a dark background')
    await Promise.resolve()
    harness
      .callbacks()
      .onToolCall({ id: 'ppt-background', name: 'set_slide_background', input: {} })
    harness.callbacks().onDone()

    await vi.waitFor(() => expect(execute).toHaveBeenCalledOnce())
    await vi.waitFor(() => expect(harness.stream).toHaveBeenCalledTimes(2))
    expect(session.snapshot().proposal).toBeUndefined()
    expect(session.snapshot().timeline.some((event) => event.kind === 'proposal')).toBe(false)
  })

  it('keeps raw Office proposals explicitly confirmation-gated in automatic PowerPoint mode', async () => {
    const harness = transportHarness()
    const proposals = createStructuredProposalController()
    const execute = vi.fn(async () => undefined)
    const session = createOfficeAgentSession({
      transport: harness.transport,
      skill: {
        id: 'powerpoint',
        systemPrompt: 'test',
        tools: [
          {
            name: 'propose_raw_office_edit',
            description: 'raw write',
            inputSchema: { type: 'object' },
          },
        ],
        executeTool: vi.fn(() => {
          const proposal = proposals.propose({
            operation: 'propose_raw_office_edit',
            toolName: 'propose_raw_office_edit',
            title: 'Raw Office edit',
            preview: {},
            impact: { host: 'powerpoint', targets: ['slide-1'], count: 1 },
            fingerprint: 'raw-v1',
            validate: async () => true,
            execute,
          })
          return {
            output: JSON.stringify({ proposalId: proposal.id }),
            mutated: false,
            summary: 'Prepared raw change',
          }
        }),
      },
      proposals,
      automaticPowerPointMutations: true,
    })

    session.send('run raw edit')
    await Promise.resolve()
    harness.callbacks().onToolCall({
      id: 'raw-write',
      name: 'propose_raw_office_edit',
      input: {},
    })
    harness.callbacks().onDone()

    await vi.waitFor(() => expect(session.snapshot().proposal).toBeDefined())
    expect(execute).not.toHaveBeenCalled()
    expect(harness.stream).toHaveBeenCalledOnce()
  })

  it.each([
    ['word', 'write_document'],
    ['excel', 'set_cell_range'],
    ['powerpoint', 'edit_slide_text'],
  ] as const)(
    'returns the confirmed %s mutation to the same AgentLoop tool call',
    async (host, toolName) => {
      const harness = transportHarness()
      const proposals = createStructuredProposalController()
      const execute = vi.fn(async () => undefined)
      const session = createOfficeAgentSession({
        transport: harness.transport,
        skill: {
          id: host,
          systemPrompt: 'test',
          tools: [{ name: toolName, description: 'write', inputSchema: { type: 'object' } }],
          executeTool: vi.fn(() => {
            const proposal = proposals.propose({
              operation: toolName,
              toolName,
              title: `Confirm ${toolName}`,
              preview: { operation: toolName },
              impact: { host, targets: ['target-1'], count: 1 },
              fingerprint: 'v1',
              validate: async () => true,
              execute,
            })
            return {
              output: JSON.stringify({
                proposalId: proposal.id,
                status: 'awaiting_user_confirmation',
              }),
              mutated: false,
              summary: 'Awaiting confirmation',
            }
          }),
        },
        proposals,
      })

      session.send(`change ${host}`)
      await Promise.resolve()
      harness.callbacks().onToolCall({ id: `${host}-write`, name: toolName, input: {} })
      harness.callbacks().onDone()
      await vi.waitFor(() => expect(session.snapshot().proposal).toBeDefined())
      const id = session.snapshot().proposal!.id

      await session.confirm(id)
      await vi.waitFor(() => expect(harness.stream).toHaveBeenCalledTimes(2))

      expect(execute).toHaveBeenCalledOnce()
      const resumed = harness.stream.mock.calls[1]?.[0] as {
        messages: Array<{ role: string; results?: Array<{ output: string }> }>
      }
      expect(resumed.messages.at(-1)).toMatchObject({
        role: 'tool',
        results: [
          {
            output: JSON.stringify({ proposalId: id, status: 'applied' }),
          },
        ],
      })
    },
  )

  it('atomically resets history and proposals when authentication is lost', async () => {
    const harness = transportHarness()
    const proposals = proposalsHarness()
    proposals.setPending()
    const session = createOfficeAgentSession({
      transport: harness.transport,
      skill: { id: 'test', systemPrompt: 'test', tools: [], executeTool: vi.fn() },
      proposals: proposals.controller,
    })
    let authLoss: (() => void) | undefined
    const signedOut = vi.fn()
    const disconnect = bindAuthLoss(
      { subscribeAuthLoss: (listener) => ((authLoss = listener), () => (authLoss = undefined)) },
      session,
      signedOut,
    )

    session.send('active request')
    await Promise.resolve()
    authLoss?.()

    expect(harness.cancel).toHaveBeenCalledOnce()
    expect(proposals.controller.logout).toHaveBeenCalledOnce()
    expect(session.snapshot()).toMatchObject({
      busy: false,
      proposal: undefined,
      assistantText: '',
    })
    expect(signedOut).toHaveBeenCalledOnce()
    disconnect()
    expect(authLoss).toBeUndefined()
  })

  it('cancels a deferred isolated visual review when reconciliation is aborted', async () => {
    const harness = transportHarness()
    let reviewer: OfficePowerPointVisualReviewer | undefined
    createOfficeAgentSession({
      transport: harness.transport,
      skill: {
        id: 'powerpoint-review',
        systemPrompt: 'test',
        tools: [],
        executeTool: vi.fn(),
        presentation: {
          prepare: () => ({ kind: 'bypass' }),
          complete: vi.fn(),
          setReviewer: (value: OfficePowerPointVisualReviewer) => {
            reviewer = value
          },
        } as never,
      },
      proposals: proposalsHarness().controller,
    })
    const controller = new AbortController()
    const pending = reviewer!.review({
      facts: {} as never,
      images: [],
      isolation: { tools: [], maxTurns: 1 },
      signal: controller.signal,
    })
    controller.abort()
    await expect(pending).resolves.toMatchObject({ status: 'cannot_verify' })
    expect(harness.cancel).toHaveBeenCalledOnce()
  })
})
