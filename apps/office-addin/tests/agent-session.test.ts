import {
  type AgentStreamCallbacks,
  type AgentTransport,
  type ToolExecution,
  suspendToolExecution,
} from '@wiswork/agent-core'
import { describe, expect, it, vi } from 'vitest'
import {
  bindAuthLoss,
  createOfficeAgentSession,
  presentationClarificationText,
} from '../src/agent/use-office-agent.js'
import {
  type ProposalDecision,
  type StructuredProposal,
  createStructuredProposalController,
} from '../src/agent/proposal-controller.js'
import { type OfficePowerPointVisualReviewer } from '../src/skills/powerpoint/powerpoint-verification.js'
import { createPcBridgeAgentTransport, type OfficeToolActivity } from '../src/agent/transport.js'
import { createOfficeDiagnostics } from '../src/diagnostics/office-diagnostics.js'
import { readFileSync } from 'node:fs'

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
    expect(diagnostics.snapshot().events.filter((event) => event.phase === 'tool')).toEqual([
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
      expect(diagnostics.snapshot().events.filter((event) => event.phase === 'tool')).toEqual([
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
      expect(diagnostics.snapshot().events.filter((event) => event.phase === 'tool')).toEqual([
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

  it.each([
    ['office_read_failed', 'list_slide_shapes', 'office_read_failed'],
    ['office_screenshot_unavailable', 'screenshot_slide', 'office_screenshot_unavailable'],
    [
      'office_screenshot_unavailable',
      'screenshot_slide',
      JSON.stringify({ error: 'office_screenshot_unavailable' }),
    ],
    [
      'office_screenshot_unavailable',
      'screenshot_slide',
      JSON.stringify({
        error: 'office_read_failed',
        reason: 'office_screenshot_unavailable',
        visualAvailableToModel: false,
      }),
    ],
  ] as const)(
    'records paired Enhanced %s failures in Taskpane diagnostics (%s, %s)',
    async (errorCode, toolName, output) => {
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
          tools: [{ name: toolName!, description: 'read', inputSchema: { type: 'object' } }],
          executeTool: vi.fn(async () => ({
            output,
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
        toolName,
        input: { slide_index: 0 },
        signal: new AbortController().signal,
      })
      expect(diagnostics.setTool).toHaveBeenCalledWith(toolName)
      expect(diagnostics.record).toHaveBeenCalledWith(
        expect.objectContaining({ phase: 'tool', errorCode }),
      )
    },
  )

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

  it('automatically resumes a transient provider failure from the current document state', async () => {
    vi.useFakeTimers()
    const harness = transportHarness()
    const session = createOfficeAgentSession({
      transport: harness.transport,
      skill: { id: 'test', systemPrompt: 'test', tools: [], executeTool: vi.fn() },
      proposals: proposalsHarness().controller,
    })

    session.send('Try this')
    await Promise.resolve()
    harness.callbacks().onError('provider_unavailable')
    expect(session.snapshot()).toMatchObject({
      busy: true,
      status: 'working',
      retryable: false,
      activity: 'Connection interrupted. Progress saved; recovering…',
    })
    expect(session.snapshot().error).toBeUndefined()

    await vi.advanceTimersByTimeAsync(2_000)

    expect(session.snapshot().status).toBe('working')
    expect(harness.stream).toHaveBeenCalledTimes(2)
    expect(
      session
        .snapshot()
        .timeline.filter((event) => event.kind === 'user')
        .map((event) => (event.kind === 'user' ? event.text : '')),
    ).toEqual(['Try this'])
    expect(harness.stream.mock.calls[1]?.[0]).toMatchObject({
      messages: expect.arrayContaining([
        expect.objectContaining({
          role: 'user',
          text: expect.stringContaining('Resume the interrupted task'),
        }),
      ]),
    })
    session.dispose()
    vi.useRealTimers()
  })

  it('shows one recoverable error only after the automatic recovery budget is exhausted', async () => {
    vi.useFakeTimers()
    const harness = transportHarness()
    const session = createOfficeAgentSession({
      transport: harness.transport,
      skill: { id: 'test', systemPrompt: 'test', tools: [], executeTool: vi.fn() },
      proposals: proposalsHarness().controller,
    })

    session.send('Build the deck')
    await Promise.resolve()
    harness.callbacks().onError('provider_unavailable')
    await vi.advanceTimersByTimeAsync(2_000)
    harness.callbacks().onError('provider_unavailable')
    await vi.advanceTimersByTimeAsync(8_000)
    harness.callbacks().onError('provider_unavailable')

    expect(harness.stream).toHaveBeenCalledTimes(3)
    expect(session.snapshot()).toMatchObject({
      busy: false,
      status: 'error',
      error: 'provider_unavailable',
      retryable: true,
      errorMessage: 'The Agent service is temporarily unavailable. Your progress is saved.',
    })
    expect(session.snapshot().timeline.filter((event) => event.kind === 'error')).toHaveLength(1)
    session.retry()
    await Promise.resolve()
    expect(harness.stream).toHaveBeenCalledTimes(4)
    expect(session.snapshot()).toMatchObject({ busy: true, status: 'working', error: undefined })
    expect(session.snapshot().timeline.filter((event) => event.kind === 'user')).toHaveLength(1)
    session.dispose()
    vi.useRealTimers()
  })

  it('cancels a scheduled automatic recovery when the user stops the task', async () => {
    vi.useFakeTimers()
    const harness = transportHarness()
    const session = createOfficeAgentSession({
      transport: harness.transport,
      skill: { id: 'test', systemPrompt: 'test', tools: [], executeTool: vi.fn() },
      proposals: proposalsHarness().controller,
    })

    session.send('Build the deck')
    await Promise.resolve()
    harness.callbacks().onError('network_error')
    session.stop()
    await vi.advanceTimersByTimeAsync(10_000)

    expect(harness.stream).toHaveBeenCalledOnce()
    expect(session.snapshot()).toMatchObject({ busy: false, status: 'cancelled' })
    session.dispose()
    vi.useRealTimers()
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

  it('automatically recovers from the bounded transport deadline', async () => {
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
      busy: true,
      status: 'working',
      activity: 'Connection interrupted. Progress saved; recovering…',
      retryable: false,
    })
    session.dispose()
  })

  it.each([
    ['session_expired', 'session_expired'],
    ['transport_auth', 'auth_required'],
    ['transport_http_401', 'auth_required'],
    ['transport_network', 'network_error'],
    ['transport_http_502', 'provider_unavailable'],
    ['transport_stream_budget_exceeded', 'transport_stream_budget_exceeded'],
  ])(
    'reports %s without attributing the run failure to the last screenshot',
    async (code, expected) => {
      const harness = transportHarness()
      const diagnostics = createOfficeDiagnostics({ host: 'powerpoint', build: 'test' })
      const session = createOfficeAgentSession({
        transport: harness.transport,
        skill: { id: 'test', systemPrompt: 'test', tools: [], executeTool: vi.fn() },
        proposals: proposalsHarness().controller,
        diagnostics,
      })
      session.send('Build a presentation')
      await Promise.resolve()
      diagnostics.setTool('screenshot_slide')
      harness.callbacks().onError(code)
      expect(session.snapshot()).toMatchObject(
        ['network_error', 'provider_unavailable'].includes(expected)
          ? { busy: true, status: 'working', error: undefined }
          : { busy: false, status: 'error', error: expected },
      )
      expect(diagnostics.snapshot().events.at(-1)).toMatchObject({
        tool: 'agent_run',
        phase: 'transport',
        error_code: expected,
      })
      session.dispose()
    },
  )

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

  it('requires explicit approval for checkpoint-backed PowerPoint proposals in automatic mode', async () => {
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
      runCheckpoint: {
        interrupted: false,
        begin: vi.fn(async () => undefined),
        finish: vi.fn(async () => undefined),
        tool: vi.fn(async () => undefined),
      },
    })

    session.send('update the slide')
    await vi.waitFor(() => expect(harness.stream).toHaveBeenCalledOnce())
    harness.callbacks().onToolCall({ id: 'ppt-write', name: 'edit_slide_text', input: {} })
    harness.callbacks().onDone()

    await vi.waitFor(() => expect(session.snapshot().proposal).toBeDefined())
    expect(execute).not.toHaveBeenCalled()
    expect(harness.stream).toHaveBeenCalledOnce()
    await session.confirm(session.snapshot().proposal!.id)
    await vi.waitFor(() => expect(execute).toHaveBeenCalledOnce())
    await vi.waitFor(() => expect(harness.stream).toHaveBeenCalledTimes(2))
    expect(session.snapshot().proposal).toBeUndefined()
    session.dispose()
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

describe('durable PPT regression coverage', () => {
  it('updates the existing approval event when an asynchronous lock review becomes ready', async () => {
    const harness = transportHarness()
    let finish!: (value: { state: 'ready'; token: string; pages: [] }) => void
    const proposals = createStructuredProposalController(undefined, {
      review: () =>
        new Promise((resolve) => {
          finish = resolve
        }),
      beforeWrite: async () => {},
      afterWrite: () => {},
    })
    const session = createOfficeAgentSession({
      transport: harness.transport,
      proposals,
      skill: {
        id: 'test',
        systemPrompt: '',
        tools: [],
        executeTool: async () => ({ output: '', summary: 'Read only' }),
      },
    })
    try {
      proposals.propose({
        operation: 'edit',
        title: 'Edit',
        preview: {},
        impact: { host: 'powerpoint', count: 1, targets: ['host'] },
        fingerprint: 'fp',
        validate: () => true,
        execute: () => {},
      })
      const event = session.snapshot().timeline.find((item) => item.kind === 'proposal')!
      expect(event).toMatchObject({ proposal: { lockReview: { state: 'checking' } } })
      await Promise.resolve()
      finish({ state: 'ready', token: 'fresh', pages: [] })
      await vi.waitFor(() =>
        expect(session.snapshot().timeline.find((item) => item.id === event.id)).toMatchObject({
          proposal: { lockReview: { state: 'ready' } },
        }),
      )
      expect(session.snapshot().timeline.filter((item) => item.kind === 'proposal')).toHaveLength(1)
    } finally {
      session.dispose()
    }
  })

  it.each(['word', 'excel', 'powerpoint'] as const)(
    'uses host-appropriate ACP activity for %s',
    async (host) => {
      const harness = transportHarness()
      const session = createOfficeAgentSession({
        host,
        transport: harness.transport,
        skill: {
          id: 'test',
          systemPrompt: 'test',
          tools: [
            { name: 'execute_office_js', description: 'execute', inputSchema: { type: 'object' } },
          ],
          executeTool: async () => ({ output: 'ok', summary: 'Executed' }),
        },
        proposals: proposalsHarness().controller,
      })
      session.send('execute')
      await vi.waitFor(() => expect(harness.stream).toHaveBeenCalledOnce())
      harness.callbacks().onToolCall({ id: 'execute-1', name: 'execute_office_js', input: {} })
      harness.callbacks().onDone()
      await vi.waitFor(() =>
        expect(session.snapshot().timeline.find((event) => event.kind === 'tool')).toMatchObject({
          summary: host === 'powerpoint' ? '页面修改操作已结束' : '已准备修改',
          state: 'complete',
        }),
      )
    },
  )

  it('does not replace another safe run that appears during document validation', async () => {
    const harness = transportHarness()
    let record = {
      runId: 'run-a',
      instruction: 'Read A',
      phase: 'running' as const,
      restartSafe: true,
    }
    const begin = vi.fn(async () => undefined)
    const session = createOfficeAgentSession({
      transport: harness.transport,
      skill: { id: 'test', systemPrompt: 'test', tools: [], executeTool: vi.fn() },
      proposals: proposalsHarness().controller,
      runCheckpoint: {
        interrupted: true,
        recovery: record,
        readRecovery: () => record,
        validateDocument: async () => {
          record = { ...record, runId: 'run-b', instruction: 'Read B' }
          return true
        },
        begin,
        finish: vi.fn(async () => undefined),
      },
    })
    await session.resumeInterrupted?.()
    expect(begin).not.toHaveBeenCalled()
    expect(harness.stream).not.toHaveBeenCalled()
  })

  it('retries a failed checkpoint begin after document validation without a recovery record', async () => {
    const harness = transportHarness()
    const begin = vi.fn(async () => undefined)
    begin.mockRejectedValueOnce(new Error('storage unavailable'))
    const validateDocument = vi.fn(async () => true)
    const session = createOfficeAgentSession({
      transport: harness.transport,
      skill: { id: 'test', systemPrompt: 'test', tools: [], executeTool: vi.fn() },
      proposals: proposalsHarness().controller,
      runCheckpoint: {
        interrupted: false,
        begin,
        finish: vi.fn(async () => undefined),
        readRecovery: () => undefined,
        validateDocument,
      },
    })
    session.send('read deck')
    await vi.waitFor(() =>
      expect(session.snapshot().error).toBe('presentation_run_checkpoint_unavailable'),
    )
    expect(harness.stream).not.toHaveBeenCalled()
    session.retry()
    await vi.waitFor(() => expect(harness.stream).toHaveBeenCalledOnce())
    expect(validateDocument).toHaveBeenCalledOnce()
    expect(begin).toHaveBeenCalledTimes(2)
  })

  it.each(['stop', 'newTask', 'logout'] as const)(
    'does not restart transient recovery after %s during validation',
    async (action) => {
      const harness = transportHarness()
      let release!: (value: boolean) => void
      const validateDocument = vi.fn(
        () =>
          new Promise<boolean>((resolve) => {
            release = resolve
          }),
      )
      const session = createOfficeAgentSession({
        transport: harness.transport,
        skill: { id: 'test', systemPrompt: 'test', tools: [], executeTool: vi.fn() },
        proposals: proposalsHarness().controller,
        runCheckpoint: {
          interrupted: false,
          begin: vi.fn(async () => undefined),
          finish: vi.fn(async () => undefined),
          validateDocument,
        },
      })
      session.send('read deck')
      await vi.waitFor(() => expect(harness.stream).toHaveBeenCalledOnce())
      harness.callbacks().onError('network_error')
      session.retry()
      session.retry()
      expect(validateDocument).toHaveBeenCalledOnce()
      session[action]()
      release(true)
      await new Promise((resolve) => setTimeout(resolve, 0))
      expect(harness.stream).toHaveBeenCalledOnce()
    },
  )

  it.each(['missing', 'throws'] as const)(
    'fails closed when latest checkpoint %s despite safe static snapshot',
    async (mode) => {
      const harness = transportHarness()
      const session = createOfficeAgentSession({
        transport: harness.transport,
        skill: { id: 'test', systemPrompt: 'test', tools: [], executeTool: vi.fn() },
        proposals: proposalsHarness().controller,
        runCheckpoint: {
          interrupted: false,
          recovery: { instruction: 'old', phase: 'running', restartSafe: true },
          readRecovery: () => {
            if (mode === 'throws') throw new Error('corrupt')
            return undefined
          },
          begin: vi.fn(async () => undefined),
          finish: vi.fn(async () => undefined),
          validateDocument: async () => true,
        },
      })
      session.send('read deck')
      await vi.waitFor(() => expect(harness.stream).toHaveBeenCalledOnce())
      harness.callbacks().onError('request_timeout')
      expect(session.snapshot().retryable).toBe(false)
      session.retry()
      expect(harness.stream).toHaveBeenCalledOnce()
    },
  )

  it('preserves a transient-failure checkpoint and blocks replay after an unsafe tool', async () => {
    const harness = transportHarness()
    const finish = vi.fn(async () => undefined)
    const session = createOfficeAgentSession({
      transport: harness.transport,
      skill: {
        id: 'test',
        systemPrompt: 'test',
        tools: [{ name: 'write', description: 'write', inputSchema: { type: 'object' } }],
        executeTool: vi.fn(async () => ({ output: 'written', mutated: true, summary: 'Written' })),
      },
      proposals: proposalsHarness().controller,
      runCheckpoint: {
        interrupted: false,
        begin: vi.fn(async () => undefined),
        tool: vi.fn(async () => undefined),
        finish,
      },
    })
    session.send('write deck')
    await vi.waitFor(() => expect(harness.stream).toHaveBeenCalledOnce())
    harness.callbacks().onToolCall({ id: 'write-1', name: 'write', input: {} })
    harness.callbacks().onDone()
    await vi.waitFor(() => expect(harness.stream).toHaveBeenCalledTimes(2))
    harness.callbacks().onError('network_error')
    expect(finish).not.toHaveBeenCalled()
    expect(session.snapshot().retryable).toBe(false)
    session.retry()
    expect(harness.stream).toHaveBeenCalledTimes(2)
  })

  it('opens after legacy prompt scrub failure but disables recovery and warns', async () => {
    const harness = transportHarness()
    const session = createOfficeAgentSession({
      transport: harness.transport,
      skill: { id: 'test', systemPrompt: 'test', tools: [], executeTool: vi.fn() },
      proposals: proposalsHarness().controller,
      runCheckpoint: {
        interrupted: true,
        scrubFailed: true,
        validateDocument: vi.fn(async () => true),
        begin: vi.fn(async () => undefined),
        finish: vi.fn(async () => undefined),
      },
    })
    expect(session.snapshot().recoveryAvailable).toBe(false)
    expect(session.snapshot().timeline[0]).toMatchObject({
      kind: 'system',
      text: expect.stringContaining('请求原文仍保留在本 PPTX'),
    })
    await session.resumeInterrupted?.()
    expect(harness.stream).not.toHaveBeenCalled()
  })

  it('resumes a pre-tool run only on explicit action after document validation', async () => {
    const harness = transportHarness()
    const validateDocument = vi.fn(async () => true)
    const session = createOfficeAgentSession({
      transport: harness.transport,
      skill: { id: 'test', systemPrompt: 'test', tools: [], executeTool: vi.fn() },
      proposals: proposalsHarness().controller,
      runCheckpoint: {
        interrupted: true,
        recovery: { instruction: 'Create deck', phase: 'running' },
        validateDocument,
        begin: vi.fn(async () => undefined),
        finish: vi.fn(async () => undefined),
      },
    })
    expect(session.snapshot().recoveryAvailable).toBe(true)
    expect(harness.stream).not.toHaveBeenCalled()
    await session.resumeInterrupted?.()
    await vi.waitFor(() => expect(harness.stream).toHaveBeenCalledOnce())
    expect(validateDocument).toHaveBeenCalledOnce()
  })

  it('never streams a recovered request when the deck switches after validation', async () => {
    const harness = transportHarness()
    let activeDocument = 'original'
    const session = createOfficeAgentSession({
      transport: harness.transport,
      skill: { id: 'test', systemPrompt: 'test', tools: [], executeTool: vi.fn() },
      proposals: proposalsHarness().controller,
      runCheckpoint: {
        interrupted: true,
        recovery: { instruction: 'Create deck', phase: 'running' },
        validateDocument: async () => {
          activeDocument = 'copy'
          return true
        },
        begin: async () => {
          if (activeDocument !== 'original') throw new Error('presentation_document_changed')
        },
        finish: vi.fn(async () => undefined),
      },
    })
    await session.resumeInterrupted?.()
    await vi.waitFor(() => expect(session.snapshot().status).toBe('error'))
    expect(harness.stream).not.toHaveBeenCalled()
  })

  it('does not resume a run after a tool boundary', async () => {
    const harness = transportHarness()
    const session = createOfficeAgentSession({
      transport: harness.transport,
      skill: { id: 'test', systemPrompt: 'test', tools: [], executeTool: vi.fn() },
      proposals: proposalsHarness().controller,
      runCheckpoint: {
        interrupted: true,
        recovery: { instruction: 'Create deck', phase: 'tool_pending', toolName: 'write_page' },
        validateDocument: vi.fn(async () => true),
        begin: vi.fn(async () => undefined),
        finish: vi.fn(async () => undefined),
      },
    })
    expect(session.snapshot().recoveryAvailable).toBe(false)
    await session.resumeInterrupted?.()
    expect(harness.stream).not.toHaveBeenCalled()
  })

  it.each([
    ['complete', '导入回执：2/2 页已记录导入'],
    ['partial', '导入回执：1/2 页已记录导入'],
    ['uncertain', '下一页写入结果不确定'],
  ] as const)(
    'shows the matched %s import receipt without replaying the tool',
    (status, detail) => {
      const harness = transportHarness()
      const session = createOfficeAgentSession({
        transport: harness.transport,
        skill: { id: 'test', systemPrompt: 'test', tools: [], executeTool: vi.fn() },
        proposals: proposalsHarness().controller,
        runCheckpoint: {
          interrupted: true,
          recovery: {
            instruction: '',
            phase: 'tool_pending',
            toolName: 'import_presentation_production',
            toolCallId: 'call-1',
            restartSafe: false,
            importReceipt: { state: status, completed: status === 'complete' ? 2 : 1, total: 2 },
          },
          begin: vi.fn(async () => undefined),
          finish: vi.fn(async () => undefined),
        },
      })
      expect(session.snapshot().recoveryAvailable).toBe(false)
      expect(session.snapshot().timeline[0]).toMatchObject({
        kind: 'system',
        text: expect.stringContaining(detail),
      })
      if (status === 'uncertain')
        expect(session.snapshot().timeline[0]).toMatchObject({
          text: expect.stringContaining('核对宿主页面'),
        })
      expect(harness.stream).not.toHaveBeenCalled()
    },
  )

  it('shows an interrupted change savepoint without offering write replay', () => {
    const harness = transportHarness()
    const session = createOfficeAgentSession({
      transport: harness.transport,
      skill: { id: 'test', systemPrompt: 'test', tools: [], executeTool: vi.fn() },
      proposals: proposalsHarness().controller,
      runCheckpoint: {
        interrupted: true,
        recovery: {
          instruction: '',
          phase: 'tool_pending',
          toolName: 'edit_existing_presentation_text',
          toolCallId: 'call-1',
          restartSafe: false,
          changeReceipt: { total: 2, unresolved: 1 },
        },
        begin: vi.fn(async () => undefined),
        finish: vi.fn(async () => undefined),
      },
    })
    expect(session.snapshot().recoveryAvailable).toBe(false)
    expect(session.snapshot().timeline[0]).toMatchObject({
      kind: 'system',
      text: expect.stringContaining('对应修改历史 2 项，其中 1 项未结算'),
    })
    expect(harness.stream).not.toHaveBeenCalled()
  })

  it('restarts an interrupted read-only run only on explicit action after document validation', async () => {
    const harness = transportHarness()
    const validateDocument = vi.fn(async () => true)
    const session = createOfficeAgentSession({
      transport: harness.transport,
      skill: { id: 'test', systemPrompt: 'test', tools: [], executeTool: vi.fn() },
      proposals: proposalsHarness().controller,
      runCheckpoint: {
        interrupted: true,
        recovery: {
          instruction: 'Inspect this presentation',
          phase: 'tool_completed',
          toolName: 'read_presentation_plan',
          restartSafe: true,
        },
        validateDocument,
        begin: vi.fn(async () => undefined),
        finish: vi.fn(async () => undefined),
      },
    })
    expect(session.snapshot().recoveryAvailable).toBe(true)
    expect(harness.stream).not.toHaveBeenCalled()
    await session.resumeInterrupted?.()
    await vi.waitFor(() => expect(harness.stream).toHaveBeenCalledOnce())
    expect(validateDocument).toHaveBeenCalledOnce()
  })

  it.each(['newTask', 'logout'] as const)(
    'does not restart an interrupted request after %s during document validation',
    async (action) => {
      const harness = transportHarness()
      let releaseValidation!: (valid: boolean) => void
      const begin = vi.fn(async () => undefined)
      const validateDocument = vi.fn(
        () =>
          new Promise<boolean>((resolve) => {
            releaseValidation = resolve
          }),
      )
      const session = createOfficeAgentSession({
        transport: harness.transport,
        skill: { id: 'test', systemPrompt: 'test', tools: [], executeTool: vi.fn() },
        proposals: proposalsHarness().controller,
        runCheckpoint: {
          interrupted: true,
          recovery: {
            instruction: 'Inspect this presentation',
            phase: 'tool_completed',
            toolName: 'read_presentation_plan',
            restartSafe: true,
          },
          validateDocument,
          begin,
          finish: vi.fn(async () => undefined),
        },
      })
      const pending = session.resumeInterrupted?.()
      expect(validateDocument).toHaveBeenCalledOnce()
      session[action]()
      releaseValidation(true)
      await pending
      expect(begin).not.toHaveBeenCalled()
      expect(harness.stream).not.toHaveBeenCalled()
    },
  )

  it('saves a checkpoint before starting and clears it after completion', async () => {
    const harness = transportHarness()
    const begin = vi.fn(async (_runId: string) => undefined)
    const finish = vi.fn(async (_runId: string) => undefined)
    const session = createOfficeAgentSession({
      transport: harness.transport,
      skill: { id: 'test', systemPrompt: 'test', tools: [], executeTool: vi.fn() },
      proposals: proposalsHarness().controller,
      runCheckpoint: { interrupted: true, begin, finish },
    })
    expect(session.snapshot().timeline[0]).toMatchObject({ kind: 'system' })
    session.send('continue')
    expect(harness.stream).not.toHaveBeenCalled()
    await vi.waitFor(() => expect(harness.stream).toHaveBeenCalledOnce())
    expect(begin).toHaveBeenCalledOnce()
    harness.callbacks().onDone()
    await vi.waitFor(() => expect(finish).toHaveBeenCalledOnce())
    expect(finish.mock.calls[0]?.[0]).toBe(begin.mock.calls[0]?.[0])
  })

  it('records a content-free run completion after a successful stream', async () => {
    const harness = transportHarness()
    const diagnostics = {
      startTrace: vi.fn(() => 'trace'),
      setTool: vi.fn(),
      record: vi.fn(),
      clear: vi.fn(),
    }
    const session = createOfficeAgentSession({
      transport: harness.transport,
      skill: { id: 'test', systemPrompt: 'test', tools: [], executeTool: vi.fn() },
      proposals: proposalsHarness().controller,
      diagnostics,
    })
    session.send('private presentation brief')
    await vi.waitFor(() => expect(harness.stream).toHaveBeenCalledOnce())
    harness.callbacks().onDone()
    await vi.waitFor(() =>
      expect(diagnostics.record).toHaveBeenCalledWith({
        phase: 'run',
        errorCode: 'agent_run_completed',
        durationMs: expect.any(Number),
      }),
    )
    expect(diagnostics.setTool).toHaveBeenCalledWith('agent_run')
    expect(JSON.stringify(diagnostics.record.mock.calls)).not.toContain(
      'private presentation brief',
    )
  })

  it('links a PowerPoint tool call to local diagnostic project and page IDs', async () => {
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
        tools: [
          {
            name: 'run_presentation_production',
            description: 'run',
            inputSchema: { type: 'object' },
          },
        ],
        executeTool: vi.fn(async () => ({
          output: 'office_write_failed',
          isError: true,
          summary: 'failed',
        })),
      },
      proposals: proposalsHarness().controller,
      diagnostics,
    })
    session.send('private brief')
    await vi.waitFor(() => expect(harness.stream).toHaveBeenCalledOnce())
    harness.callbacks().onToolCall({
      id: 'call-1',
      name: 'run_presentation_production',
      input: {
        project_id: 'project-1',
        request_id: 'run-1',
        page_id: 'page-3',
        private_text: 'secret',
      },
    })
    harness.callbacks().onDone()
    await vi.waitFor(() =>
      expect(diagnostics.setTool).toHaveBeenCalledWith('run_presentation_production', {
        project_id: 'project-1',
        request_id: 'run-1',
        page_id: 'page-3',
        tool_call_id: 'call-1',
      }),
    )
    expect(JSON.stringify(diagnostics.setTool.mock.calls)).not.toContain('secret')
  })

  it('waits for an ordinary tool completion checkpoint before the next request', async () => {
    const harness = transportHarness()
    let savePending!: () => void
    let saveCompleted!: () => void
    const executeTool = vi.fn(async () => ({ output: 'read', summary: 'Read document' }))
    const begin = vi.fn(async (_runId: string) => undefined)
    const tool = vi.fn(
      (_runId: string, phase: 'tool_pending' | 'tool_completed') =>
        new Promise<void>((resolve) => {
          if (phase === 'tool_pending') savePending = resolve
          else saveCompleted = resolve
        }),
    )
    const session = createOfficeAgentSession({
      transport: harness.transport,
      skill: {
        id: 'test',
        systemPrompt: 'test',
        tools: [{ name: 'read_document', description: 'read', inputSchema: { type: 'object' } }],
        executeTool,
      },
      proposals: proposalsHarness().controller,
      runCheckpoint: { interrupted: false, begin, tool, finish: vi.fn(async () => undefined) },
    })

    session.send('read')
    await vi.waitFor(() => expect(harness.stream).toHaveBeenCalledOnce())
    harness.callbacks().onToolCall({ id: 'read-1', name: 'read_document', input: {} })
    harness.callbacks().onDone()
    await vi.waitFor(() =>
      expect(tool).toHaveBeenCalledWith(
        expect.any(String),
        'tool_pending',
        'read_document',
        false,
        'read-1',
      ),
    )
    expect(executeTool).not.toHaveBeenCalled()

    savePending()
    await vi.waitFor(() =>
      expect(tool).toHaveBeenCalledWith(
        expect.any(String),
        'tool_completed',
        'read_document',
        false,
        'read-1',
      ),
    )
    expect(executeTool).toHaveBeenCalledOnce()
    expect(tool.mock.calls[0]?.[0]).toBe(begin.mock.calls[0]?.[0])
    expect(tool.mock.calls[1]?.[0]).toBe(begin.mock.calls[0]?.[0])
    expect(harness.stream).toHaveBeenCalledOnce()

    saveCompleted()
    await vi.waitFor(() => expect(harness.stream).toHaveBeenCalledTimes(2))
  })

  it('halts after a completed tool when its final checkpoint cannot be saved', async () => {
    const harness = transportHarness()
    const finish = vi.fn(async () => undefined)
    const executeTool = vi.fn(async () => ({
      output: 'changed',
      summary: 'Changed',
      mutated: true,
    }))
    const session = createOfficeAgentSession({
      transport: harness.transport,
      skill: {
        id: 'test',
        systemPrompt: 'test',
        tools: [{ name: 'write', description: 'write', inputSchema: { type: 'object' } }],
        executeTool,
      },
      proposals: proposalsHarness().controller,
      runCheckpoint: {
        interrupted: false,
        begin: vi.fn(async () => undefined),
        tool: vi.fn(async (_id, phase) => {
          if (phase === 'tool_completed') throw new Error('save failed')
        }),
        finish,
      },
    })
    session.send('change')
    await vi.waitFor(() => expect(harness.stream).toHaveBeenCalledOnce())
    harness.callbacks().onToolCall({ id: 'write-1', name: 'write', input: {} })
    harness.callbacks().onDone()
    await vi.waitFor(() =>
      expect(session.snapshot().error).toBe('presentation_run_checkpoint_unavailable'),
    )
    expect(executeTool).toHaveBeenCalledOnce()
    expect(harness.stream).toHaveBeenCalledOnce()
    expect(finish).not.toHaveBeenCalled()
    expect(session.snapshot().retryable).toBe(false)
  })

  it.each(['newTask', 'logout'] as const)(
    'keeps the pending write checkpoint when %s resets a tool still executing',
    async (action) => {
      const harness = transportHarness()
      let finishWrite!: (result: ToolExecution) => void
      const executeTool = vi.fn(
        () =>
          new Promise<ToolExecution>((resolve) => {
            finishWrite = resolve
          }),
      )
      const begin = vi.fn(async () => undefined)
      const tool = vi.fn(async () => undefined)
      const finish = vi.fn(async () => undefined)
      const session = createOfficeAgentSession({
        transport: harness.transport,
        skill: {
          id: 'test',
          systemPrompt: 'test',
          tools: [{ name: 'write', description: 'write', inputSchema: { type: 'object' } }],
          executeTool,
        },
        proposals: proposalsHarness().controller,
        runCheckpoint: { interrupted: false, begin, tool, finish },
      })
      session.send('change')
      await vi.waitFor(() => expect(harness.stream).toHaveBeenCalledOnce())
      harness.callbacks().onToolCall({ id: 'write-1', name: 'write', input: {} })
      harness.callbacks().onDone()
      await vi.waitFor(() => expect(executeTool).toHaveBeenCalledOnce())
      expect(tool).toHaveBeenCalledWith(
        expect.any(String),
        'tool_pending',
        'write',
        false,
        'write-1',
      )

      session[action]()
      expect(finish).not.toHaveBeenCalled()
      finishWrite({ output: 'changed', summary: 'Changed', mutated: true })
      await Promise.resolve()
      expect(tool).not.toHaveBeenCalledWith(
        expect.any(String),
        'tool_completed',
        'write',
        true,
        'write-1',
      )
      expect(finish).not.toHaveBeenCalled()
    },
  )

  it('does not execute an old tool after its pending checkpoint outlives the run', async () => {
    const harness = transportHarness()
    let releasePending!: () => void
    const executeTool = vi.fn(async () => ({ output: 'done', summary: 'Done' }))
    const tool = vi.fn((_runId: string, phase: 'tool_pending' | 'tool_completed') =>
      phase === 'tool_pending'
        ? new Promise<void>((resolve) => {
            releasePending = resolve
          })
        : Promise.resolve(),
    )
    const session = createOfficeAgentSession({
      transport: harness.transport,
      skill: {
        id: 'test',
        systemPrompt: 'test',
        tools: [{ name: 'read', description: 'read', inputSchema: { type: 'object' } }],
        executeTool,
      },
      proposals: proposalsHarness().controller,
      runCheckpoint: {
        interrupted: false,
        begin: vi.fn(async () => undefined),
        tool,
        finish: vi.fn(async () => undefined),
      },
    })
    session.send('old')
    await vi.waitFor(() => expect(harness.stream).toHaveBeenCalledOnce())
    harness.callbacks().onToolCall({ id: 'old-tool', name: 'read', input: {} })
    harness.callbacks().onDone()
    await vi.waitFor(() =>
      expect(tool).toHaveBeenCalledWith(
        expect.any(String),
        'tool_pending',
        'read',
        false,
        'old-tool',
      ),
    )
    session.newTask()
    session.send('new')
    await vi.waitFor(() => expect(harness.stream).toHaveBeenCalledTimes(2))
    releasePending()
    await Promise.resolve()
    expect(executeTool).not.toHaveBeenCalled()
    expect(tool).toHaveBeenCalledTimes(1)
  })

  it('waits for proposal decision and completion checkpoint before continuing', async () => {
    const harness = transportHarness()
    const proposals = proposalsHarness()
    let releaseCompleted!: () => void
    const tool = vi.fn((_runId: string, phase: 'tool_pending' | 'tool_completed') =>
      phase === 'tool_completed'
        ? new Promise<void>((resolve) => {
            releaseCompleted = resolve
          })
        : Promise.resolve(),
    )
    const session = createOfficeAgentSession({
      transport: harness.transport,
      skill: {
        id: 'test',
        systemPrompt: 'test',
        tools: [{ name: 'propose', description: 'propose', inputSchema: { type: 'object' } }],
        executeTool: vi.fn(async () => {
          proposals.setPending()
          return { output: 'prepared', summary: 'Prepared' }
        }),
      },
      proposals: proposals.controller,
      runCheckpoint: {
        interrupted: false,
        begin: vi.fn(async () => undefined),
        tool,
        finish: vi.fn(async () => undefined),
      },
    })
    session.send('edit')
    await vi.waitFor(() => expect(harness.stream).toHaveBeenCalledOnce())
    harness.callbacks().onToolCall({ id: 'proposal-tool', name: 'propose', input: {} })
    harness.callbacks().onDone()
    await vi.waitFor(() => expect(session.snapshot().proposal?.id).toBe('p1'))
    expect(tool).toHaveBeenCalledTimes(1)
    await session.confirm('p1')
    await vi.waitFor(() =>
      expect(tool).toHaveBeenCalledWith(
        expect.any(String),
        'tool_completed',
        'propose',
        true,
        'proposal-tool',
      ),
    )
    expect(harness.stream).toHaveBeenCalledOnce()
    releaseCompleted()
    await vi.waitFor(() => expect(harness.stream).toHaveBeenCalledTimes(2))
  })

  it('halts after approved work if its completion checkpoint fails', async () => {
    const harness = transportHarness()
    const proposals = proposalsHarness()
    const finish = vi.fn(async () => undefined)
    const session = createOfficeAgentSession({
      transport: harness.transport,
      skill: {
        id: 'test',
        systemPrompt: 'test',
        tools: [{ name: 'propose', description: 'propose', inputSchema: { type: 'object' } }],
        executeTool: vi.fn(async () => {
          proposals.setPending()
          return { output: 'prepared', summary: 'Prepared' }
        }),
      },
      proposals: proposals.controller,
      runCheckpoint: {
        interrupted: false,
        begin: vi.fn(async () => undefined),
        tool: vi.fn(async (_id, phase) => {
          if (phase === 'tool_completed') throw new Error('save failed')
        }),
        finish,
      },
    })
    session.send('edit')
    await vi.waitFor(() => expect(harness.stream).toHaveBeenCalledOnce())
    harness.callbacks().onToolCall({ id: 'proposal-tool', name: 'propose', input: {} })
    harness.callbacks().onDone()
    await vi.waitFor(() => expect(session.snapshot().proposal?.id).toBe('p1'))
    await session.confirm('p1')
    await vi.waitFor(() =>
      expect(session.snapshot().error).toBe('presentation_run_checkpoint_unavailable'),
    )
    expect(harness.stream).toHaveBeenCalledOnce()
    expect(finish).not.toHaveBeenCalled()
  })

  it('waits for a suspended tool result and its completion checkpoint', async () => {
    const harness = transportHarness()
    let releaseResult!: (result: ToolExecution) => void
    let releaseCompleted!: () => void
    const tool = vi.fn((_runId: string, phase: 'tool_pending' | 'tool_completed') =>
      phase === 'tool_completed'
        ? new Promise<void>((resolve) => {
            releaseCompleted = resolve
          })
        : Promise.resolve(),
    )
    const session = createOfficeAgentSession({
      transport: harness.transport,
      skill: {
        id: 'test',
        systemPrompt: 'test',
        tools: [{ name: 'wait', description: 'wait', inputSchema: { type: 'object' } }],
        executeTool: vi.fn(async () =>
          suspendToolExecution(
            new Promise<ToolExecution>((resolve) => {
              releaseResult = resolve
            }),
          ),
        ),
      },
      proposals: proposalsHarness().controller,
      runCheckpoint: {
        interrupted: false,
        begin: vi.fn(async () => undefined),
        tool,
        finish: vi.fn(async () => undefined),
      },
    })
    session.send('wait')
    await vi.waitFor(() => expect(harness.stream).toHaveBeenCalledOnce())
    harness.callbacks().onToolCall({ id: 'wait-tool', name: 'wait', input: {} })
    harness.callbacks().onDone()
    await vi.waitFor(() => expect(releaseResult).toBeTypeOf('function'))
    expect(tool).toHaveBeenCalledTimes(1)
    releaseResult({ output: 'done', summary: 'Done' })
    await vi.waitFor(() =>
      expect(tool).toHaveBeenCalledWith(
        expect.any(String),
        'tool_completed',
        'wait',
        false,
        'wait-tool',
      ),
    )
    expect(harness.stream).toHaveBeenCalledOnce()
    releaseCompleted()
    await vi.waitFor(() => expect(harness.stream).toHaveBeenCalledTimes(2))
  })

  it('does not start a run when its checkpoint cannot be saved', async () => {
    const harness = transportHarness()
    const session = createOfficeAgentSession({
      transport: harness.transport,
      skill: { id: 'test', systemPrompt: 'test', tools: [], executeTool: vi.fn() },
      proposals: proposalsHarness().controller,
      runCheckpoint: {
        interrupted: false,
        begin: async () => {
          throw new Error('save failed')
        },
        finish: vi.fn(async () => undefined),
      },
    })
    session.send('continue')
    await vi.waitFor(() => expect(session.snapshot().status).toBe('error'))
    expect(session.snapshot().error).toBe('presentation_run_checkpoint_unavailable')
    expect(harness.stream).not.toHaveBeenCalled()
  })

  it('passes a longer presentation brief to the bounded local recovery checkpoint', async () => {
    const harness = transportHarness()
    const begin = vi.fn(async () => undefined)
    const session = createOfficeAgentSession({
      transport: harness.transport,
      skill: { id: 'test', systemPrompt: 'test', tools: [], executeTool: vi.fn() },
      proposals: proposalsHarness().controller,
      runCheckpoint: { interrupted: false, begin, finish: vi.fn(async () => undefined) },
    })
    const brief = '页面、来源和样式要求。'.repeat(150)
    expect(brief.length).toBeGreaterThan(1000)
    session.send(brief)
    await vi.waitFor(() => expect(begin).toHaveBeenCalledOnce())
    expect(begin).toHaveBeenCalledWith(expect.any(String), brief)
    session.stop()
  })

  it('does not start a cancelled run after a delayed checkpoint save', async () => {
    const harness = transportHarness()
    let release!: () => void
    const begin = vi.fn(
      () =>
        new Promise<void>((resolve) => {
          release = resolve
        }),
    )
    const finish = vi.fn(async (_runId: string) => undefined)
    const session = createOfficeAgentSession({
      transport: harness.transport,
      skill: { id: 'test', systemPrompt: 'test', tools: [], executeTool: vi.fn() },
      proposals: proposalsHarness().controller,
      runCheckpoint: { interrupted: false, begin, finish },
    })
    session.send('do work')
    session.stop()
    release()
    await vi.waitFor(() => expect(finish).toHaveBeenCalledOnce())
    expect(harness.stream).not.toHaveBeenCalled()
    expect(session.snapshot().status).toBe('cancelled')
  })

  it('shows PowerPoint production as a stage activity through ACP updates', async () => {
    const harness = transportHarness()
    let finish!: () => void
    const session = createOfficeAgentSession({
      transport: harness.transport,
      skill: {
        id: 'test',
        systemPrompt: 'test',
        tools: [
          {
            name: 'run_presentation_production',
            description: 'run',
            inputSchema: { type: 'object' },
          },
        ],
        executeTool: vi.fn(
          () =>
            new Promise<ToolExecution>((resolve) => {
              finish = () => resolve({ output: '{}', summary: 'internal', mutated: false })
            }),
        ),
      },
      proposals: proposalsHarness().controller,
    })
    session.send('继续制作')
    await Promise.resolve()
    harness
      .callbacks()
      .onToolCall({ id: 'production-1', name: 'run_presentation_production', input: {} })
    harness.callbacks().onDone()
    await vi.waitFor(() =>
      expect(session.snapshot().timeline).toEqual(
        expect.arrayContaining([
          expect.objectContaining({ kind: 'tool', summary: '正在处理逐页制作…', state: 'running' }),
        ]),
      ),
    )
    finish()
    await vi.waitFor(() =>
      expect(session.snapshot().timeline).toEqual(
        expect.arrayContaining([
          expect.objectContaining({
            kind: 'tool',
            summary: '逐页制作操作已结束',
            state: 'complete',
          }),
        ]),
      ),
    )
  })

  it('does not duplicate the user instruction while automatic recovery is pending', async () => {
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
    ).toEqual(['Try this'])
    session.dispose()
  })

  it('keeps P0-20 transient model error and user cancellation distinct from completion', async () => {
    const scenario = JSON.parse(
      readFileSync(
        new URL(
          '../../../docs/product/ppt-benchmark-materials/PPT-P0-20/scenario.json',
          import.meta.url,
        ),
        'utf8',
      ),
    )
    expect(scenario.faultSchedule.map((fault: { id: string }) => fault.id)).toEqual([
      'F1',
      'F2',
      'F3',
    ])
    const harness = transportHarness()
    const begin = vi.fn(async () => undefined)
    const finish = vi.fn(async () => undefined)
    const session = createOfficeAgentSession({
      transport: harness.transport,
      skill: { id: 'p0-20', systemPrompt: 'test', tools: [], executeTool: vi.fn() },
      proposals: proposalsHarness().controller,
      runCheckpoint: { interrupted: false, begin, finish, validateDocument: async () => true },
    })
    const instruction = '继续同一项目的八页制作'
    session.send(instruction)
    await vi.waitFor(() => expect(harness.stream).toHaveBeenCalledTimes(1))
    harness.callbacks().onError('provider_unavailable')
    expect(session.snapshot()).toMatchObject({ status: 'error', retryable: true })
    expect(
      session
        .snapshot()
        .timeline.some((event) => event.kind === 'system' && event.text.includes('完成')),
    ).toBe(false)
    session.retry()
    await vi.waitFor(() => expect(harness.stream).toHaveBeenCalledTimes(2))
    expect(begin).toHaveBeenCalledTimes(2)
    expect(
      session
        .snapshot()
        .timeline.filter((event) => event.kind === 'user')
        .map((event) => (event.kind === 'user' ? event.text : '')),
    ).toEqual([instruction, instruction])
    session.stop()
    harness.callbacks().onDone()
    await vi.waitFor(() => expect(session.snapshot().status).toBe('cancelled'))
    expect(session.snapshot()).toMatchObject({ busy: false, retryable: false })
    expect(harness.cancel).toHaveBeenCalledOnce()
    expect(finish).toHaveBeenCalled()
  })

  it('reports screenshot metadata without sending bulk images into model history', async () => {
    const harness = transportHarness()
    const proposals = createStructuredProposalController()
    const pngBase64 =
      'iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAQAAAC1HAwCAAAAC0lEQVR42mP8/x8AAwMCAO+aPioAAAAASUVORK5CYII='
    const session = createOfficeAgentSession({
      transport: harness.transport,
      skill: {
        id: 'powerpoint',
        systemPrompt: 'test',
        tools: [{ name: 'edit_slide_text', description: 'write', inputSchema: { type: 'object' } }],
        executeTool: () => {
          const proposal = proposals.propose({
            operation: 'edit_slide_text',
            title: 'Edit',
            preview: {},
            impact: { host: 'powerpoint', targets: ['slide-1'], count: 1 },
            fingerprint: 'v1',
            validate: () => true,
            execute: () => {},
            postWrite: () => ({
              status: 'captured',
              pages: [
                {
                  slideId: 'slide-1',
                  pngBase64,
                  digest: 'f4b555ad4009f54a1a37dc29e7ccf9f8f4cfe22410ba7769061c0328cdb6db67',
                },
              ],
            }),
          })
          return {
            output: JSON.stringify({ proposalId: proposal.id }),
            mutated: false,
            summary: 'Awaiting confirmation',
          }
        },
      },
      proposals,
    })
    session.send('edit')
    await Promise.resolve()
    harness.callbacks().onToolCall({ id: 'write', name: 'edit_slide_text', input: {} })
    harness.callbacks().onDone()
    await vi.waitFor(() => expect(session.snapshot().proposal).toBeDefined())
    await session.confirm(session.snapshot().proposal!.id)
    await vi.waitFor(() => expect(harness.stream).toHaveBeenCalledTimes(2))
    const resumed = harness.stream.mock.calls[1]?.[0] as {
      messages: Array<{ results?: Array<{ output: string; content?: unknown }> }>
    }
    const result = resumed.messages.at(-1)?.results?.[0]
    expect(JSON.parse(result!.output)).toMatchObject({
      status: 'applied',
      qaPassed: false,
      visualReview: 'pending',
      postWrite: { status: 'captured', pages: [{ slideId: 'slide-1' }] },
    })
    expect(result!.content).toBeUndefined()
  })

  it('attributes a run transport failure to the run rather than the last presentation page', async () => {
    const harness = transportHarness()
    const diagnostics = createOfficeDiagnostics({ host: 'powerpoint', build: 'test' })
    const session = createOfficeAgentSession({
      transport: harness.transport,
      skill: { id: 'test', systemPrompt: 'test', tools: [], executeTool: vi.fn() },
      proposals: proposalsHarness().controller,
      diagnostics,
    })

    session.send('Produce a deck')
    await Promise.resolve()
    diagnostics.setTool('run_presentation_production', {
      project_id: 'project-1',
      request_id: 'run-1',
      page_id: 'page-3',
    })
    harness.callbacks().onError('transport_timeout')

    const event = diagnostics.snapshot().events.at(-1)
    expect(event).toMatchObject({
      tool: 'agent_run',
      phase: 'transport',
      error_code: 'request_timeout',
    })
    expect(event).not.toHaveProperty('presentation_context')
    expect(event).not.toHaveProperty('presentation_stage')
  })

  it('keeps checkpoint-backed transport timeouts available for validated manual recovery', async () => {
    const harness = transportHarness()
    const session = createOfficeAgentSession({
      runCheckpoint: {
        interrupted: false,
        recovery: { instruction: 'Build a complex presentation', phase: 'running' },
        validateDocument: async () => true,
        begin: vi.fn(async () => undefined),
        finish: vi.fn(async () => undefined),
      },
      transport: harness.transport,
      skill: { id: 'test', systemPrompt: 'test', tools: [], executeTool: vi.fn() },
      proposals: proposalsHarness().controller,
    })

    session.send('Build a complex presentation')
    await vi.waitFor(() => expect(harness.stream).toHaveBeenCalledOnce())
    harness.callbacks().onError('transport_timeout')

    expect(session.snapshot()).toMatchObject({
      error: 'request_timeout',
      errorMessage: expect.stringContaining('未自动重放'),
      retryable: true,
    })
  })

  it('hides unknown backup confirmation response text from the session', async () => {
    const proposals = proposalsHarness()
    proposals.controller.confirm.mockRejectedValue(new Error('quota_exceeded /private/secret'))
    proposals.setPending()
    const session = createOfficeAgentSession({
      transport: transportHarness().transport,
      skill: { id: 'test', systemPrompt: 'test', tools: [], executeTool: vi.fn() },
      proposals: proposals.controller,
    })
    await session.confirm('p1')
    expect(session.snapshot()).toMatchObject({
      error: 'office_write_failed',
      errorMessage: 'The approved change could not be applied.',
      retryable: false,
    })
    expect(JSON.stringify(session.snapshot())).not.toContain('/private/secret')
    expect(proposals.controller.confirm).toHaveBeenCalledOnce()
  })

  it.each([
    'capture_presentation_page_qa',
    'record_presentation_page_review',
    'compare_presentation_page_structure',
  ])(
    'records nonfatal waiting screenshot diagnostics for %s without changing its model output',
    async (name) => {
      const harness = transportHarness()
      const output = JSON.stringify({
        status: 'waiting_screenshot',
        pageId: 'page-1',
        hostSlideId: 'host-1',
        retryable: true,
      })
      const diagnostics = createOfficeDiagnostics({ host: 'powerpoint', build: 'screenshot-test' })
      const session = createOfficeAgentSession({
        transport: harness.transport,
        skill: {
          id: 'qa-test',
          systemPrompt: 'test',
          tools: [{ name, description: 'test', inputSchema: { type: 'object' } }],
          executeTool: async () => ({ output, mutated: false, summary: 'waiting' }),
        },
        proposals: proposalsHarness().controller,
        diagnostics,
      })
      try {
        session.send('Capture this page')
        await Promise.resolve()
        harness.callbacks().onToolCall({
          id: 'screenshot-call',
          name,
          input: { project_id: 'project-1', page_id: 'page-1' },
        })
        harness.callbacks().onDone()
        await vi.waitFor(() =>
          expect(
            diagnostics
              .snapshot()
              .events.some((event) => event.error_code === 'presentation_screenshot_waiting'),
          ).toBe(true),
        )
        const event = diagnostics
          .snapshot()
          .events.find((event) => event.error_code === 'presentation_screenshot_waiting')!
        expect(event).toMatchObject({
          tool: name,
          phase: 'tool',
          outcome: 'unsupported',
          presentation_context: {
            project_id: 'project-1',
            page_id: 'page-1',
            tool_call_id: 'screenshot-call',
          },
        })
        expect(session.snapshot().status).not.toBe('error')
        expect(diagnostics.exportJson()).not.toContain('page-1')
        expect(JSON.stringify(harness.stream.mock.calls)).toContain('waiting_screenshot')
      } finally {
        session.dispose()
      }
    },
  )

  it('keeps a real unresolved screenshot failure distinct from a generic agent failure', async () => {
    const harness = transportHarness(),
      diagnostics = createOfficeDiagnostics({ host: 'powerpoint', build: 'screenshot-test' })
    const name = 'capture_presentation_page_qa'
    const session = createOfficeAgentSession({
      transport: harness.transport,
      skill: {
        id: 'qa-test',
        systemPrompt: 'test',
        tools: [{ name, description: 'test', inputSchema: { type: 'object' } }],
        executeTool: async () => ({
          output: 'presentation_qa_attempt_unresolved',
          isError: true,
          mutated: false,
          summary: 'unresolved',
        }),
      },
      proposals: proposalsHarness().controller,
      diagnostics,
    })
    try {
      session.send('Capture')
      await Promise.resolve()
      harness.callbacks().onToolCall({ id: 'failed-capture', name, input: { page_id: 'page-1' } })
      harness.callbacks().onDone()
      await vi.waitFor(() =>
        expect(
          diagnostics
            .snapshot()
            .events.some((event) => event.error_code === 'presentation_qa_attempt_unresolved'),
        ).toBe(true),
      )
      expect(
        diagnostics
          .snapshot()
          .events.find((event) => event.error_code === 'presentation_qa_attempt_unresolved')
          ?.outcome,
      ).toBe('failed')
    } finally {
      session.dispose()
    }
  })

  it.each([
    { name: 'unrelated_tool', payload: {}, mutated: false },
    { name: 'capture_presentation_page_qa', payload: { pageId: 'other-page' }, mutated: false },
    { name: 'capture_presentation_page_qa', payload: { retryable: false }, mutated: false },
    { name: 'capture_presentation_page_qa', payload: { hostSlideId: '' }, mutated: false },
    { name: 'capture_presentation_page_qa', payload: { private: 'unexpected' }, mutated: false },
    {
      name: 'capture_presentation_page_qa',
      payload: { hostSlideId: 'x'.repeat(4096) },
      mutated: false,
    },
    { name: 'capture_presentation_page_qa', payload: {}, mutated: true },
  ])(
    'does not infer screenshot waiting from untrusted or unrelated results: %j',
    async ({ name, payload, mutated }) => {
      const harness = transportHarness()
      const diagnostics = createOfficeDiagnostics({ host: 'powerpoint', build: 'screenshot-test' })
      const output = JSON.stringify({
        status: 'waiting_screenshot',
        pageId: 'page-1',
        hostSlideId: 'host-1',
        retryable: true,
        ...payload,
      })
      const session = createOfficeAgentSession({
        transport: harness.transport,
        skill: {
          id: 'qa-test',
          systemPrompt: 'test',
          tools: [{ name, description: 'test', inputSchema: { type: 'object' } }],
          executeTool: async () => ({ output, mutated, summary: 'done' }),
        },
        proposals: proposalsHarness().controller,
        diagnostics,
      })
      try {
        session.send('Capture this page')
        await Promise.resolve()
        harness.callbacks().onToolCall({
          id: 'screenshot-call',
          name,
          input: { project_id: 'project-1', page_id: 'page-1' },
        })
        harness.callbacks().onDone()
        await vi.waitFor(() => expect(harness.stream.mock.calls.length).toBeGreaterThan(1))
        expect(
          diagnostics
            .snapshot()
            .events.some((event) => event.error_code === 'presentation_screenshot_waiting'),
        ).toBe(false)
      } finally {
        session.dispose()
      }
    },
  )
})
