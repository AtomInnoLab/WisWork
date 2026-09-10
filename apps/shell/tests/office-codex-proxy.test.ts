import { describe, expect, it, vi } from 'vitest'
import { ENHANCED_HOSTS, type EnhancedRolloutPolicy } from '@wiswork/agent-runtime'
import {
  encodeOfficeScreenshotResult,
  isToolExecutionSuspension,
  type ToolExecution,
} from '@wiswork/agent-core'
import type { DocumentToolSession } from '@wiswork/codex-bridge'
import { OFFICE_PROXY_KEEPALIVE_MS, createOfficeCodexProxy } from '../src/main/office-codex-proxy'
import { createShellEnhancedPolicyAuthority } from '../src/main/enhanced-policy-authority'
import { createStructuredProposalController } from '../../office-addin/src/agent/proposal-controller'
import type { PowerPointAdapter } from '../../office-addin/src/skills/powerpoint/browser-powerpoint-adapter'
import { createPowerPointSkill } from '../../office-addin/src/skills/powerpoint/powerpoint-skill'

// Legacy gateway event shape (without turnId) remains supported.
async function semanticCall(input: any, call: any) {
  input.onEvent({ type: 'tool-start', callId: call.id, toolName: call.name })
  const result = await input.toolSession.callTool(input.toolSession.credentials, call)
  input.onEvent({
    type: 'tool-complete',
    callId: call.id,
    toolName: call.name,
    isError: result.isError === true,
  })
  return result
}

const rollout: EnhancedRolloutPolicy = {
  globalEnabled: true,
  rawOfficeEnabled: false,
  hosts: Object.fromEntries(ENHANCED_HOSTS.map((host) => [host, true])) as any,
}
const statement = {
  version: 1,
  runtime_mode: 'enhanced',
  runtime_instance: 'runtime_0123456789abcdef',
  component_version: '0.147.0',
  host: 'office-word',
  raw_office: false,
  expires_at: Date.now() + 60_000,
  policy_generation: 0,
  session_generation: 3,
} as const

describe('Office Codex proxy', () => {
  it.each(['office_verify_failed', 'proposal_stale'])(
    'preserves the mutation signal for a failed remote proposal: %s',
    async (error) => {
      let receipt!: ToolExecution
      const proxy = createOfficeCodexProxy({
        runtime: {
          async runOfficeTurn(input: any) {
            const result = input.toolSession.callTool(input.toolSession.credentials, {
              id: 'repair_call',
              name: 'edit_slide_text',
              input: { slide_index: 0, shape_id: 'title', text: 'Title' },
            })
            receipt = await (isToolExecutionSuspension(result) ? result.result : result)
            input.onEvent({ type: 'terminal', status: 'completed' })
          },
        } as any,
        rollout,
        policyAuthority: createShellEnhancedPolicyAuthority(() => 0),
      })
      const response = await proxy({
        body: {
          system: '',
          messages: [],
          tools: [
            { name: 'edit_slide_text', description: 'edit', input_schema: { type: 'object' } },
          ],
        },
        signal: new AbortController().signal,
        host: 'PowerPoint',
        sessionId: 'session_12345678',
        requestId: 'request_12345678',
        statement: { ...statement, host: 'office-powerpoint' },
        executeTool: async () => ({
          output: JSON.stringify({ proposalId: 'p1', status: 'failed', error }),
          isError: true,
        }),
      })
      for await (const _chunk of response.body as AsyncIterable<Uint8Array>) {
        /* drain */
      }
      expect(receipt).toMatchObject({ isError: true, mutated: error === 'office_verify_failed' })
    },
  )

  it.each([
    'success',
    'normalized',
    'normalizer-failed',
    'unavailable',
    'large',
    'injected',
    'cancelled',
    'remote-failed',
  ] as const)(
    'prefetches bounded private image bytes on PC before dispatch: %s',
    async (scenario) => {
      const input = {
        url: 'https://images.example/approved.png',
        slide_index: 0,
        left: 1,
        top: 2,
        width: 30,
        height: 40,
        ...(scenario === 'injected' ? { _wiswork_image_base64: 'model-injected' } : {}),
      }
      const bytes = Buffer.from('private-image-bytes')
      const jpegBytes = Buffer.from([0xff, 0xd8, 0xff, 0xd9])
      const prepareImageHandoff = vi.fn(async () => {
        if (scenario === 'normalizer-failed') throw new Error('image_limit')
        return { mime: 'image/jpeg' as const, bytes: jpegBytes }
      })
      let receipt!: ToolExecution
      let rejectedBeforeApproval = false
      const controller = new AbortController()
      const executeTool = vi.fn(async (_call: { input: Record<string, unknown> }) => {
        if (scenario === 'remote-failed') throw new Error('private uncertain host write')
        return { output: 'applied', isError: false }
      })
      const executeRetrieval = vi.fn(async () => {
        if (scenario === 'unavailable') throw new Error('private upstream credentials')
        if (scenario === 'cancelled') controller.abort()
        return new TextEncoder().encode(
          JSON.stringify({
            mime: 'image/png',
            data_base64: (['large', 'normalized', 'normalizer-failed'].includes(scenario)
              ? Buffer.alloc(181 * 1024)
              : bytes
            ).toString('base64'),
          }),
        )
      })
      const runtime = {
        async runOfficeTurn(value: any) {
          value.onEvent({ type: 'tool-start', callId: 'image_call', toolName: 'insert_web_image' })
          if (scenario === 'injected') {
            try {
              value.summarizeProposal({ name: 'insert_web_image', input })
            } catch (error) {
              rejectedBeforeApproval =
                error instanceof Error && error.message === 'invalid_tool_input'
            }
          }
          const result = value.toolSession.callTool(value.toolSession.credentials, {
            id: 'image_call',
            name: 'insert_web_image',
            input,
          })
          receipt = await (isToolExecutionSuspension(result) ? result.result : result)
          value.onEvent({
            type: 'tool-complete',
            callId: 'image_call',
            toolName: 'insert_web_image',
            isError: receipt.isError === true,
          })
          value.onEvent({ type: 'terminal', status: 'completed' })
        },
      }
      const proxy = createOfficeCodexProxy({
        runtime: runtime as any,
        rollout,
        ...(['normalized', 'normalizer-failed'].includes(scenario) ? { prepareImageHandoff } : {}),
        policyAuthority: createShellEnhancedPolicyAuthority(() => 0),
      })
      const response = await proxy({
        body: {
          system: '',
          messages: [],
          tools: [
            { name: 'insert_web_image', description: 'insert', input_schema: { type: 'object' } },
          ],
        },
        signal: controller.signal,
        host: 'PowerPoint',
        sessionId: 'session_12345678',
        requestId: 'request_12345678',
        statement: { ...statement, host: 'office-powerpoint' },
        executeTool,
        executeRetrieval,
      })
      let stream = ''
      for await (const chunk of response.body as AsyncIterable<Uint8Array>)
        stream += new TextDecoder().decode(chunk)
      if (scenario === 'success' || scenario === 'normalized') {
        expect(executeRetrieval).toHaveBeenCalledWith(
          'image-fetch.v1',
          { url: input.url },
          expect.any(AbortSignal),
        )
        expect(executeTool).toHaveBeenCalledWith(
          expect.objectContaining({
            toolName: 'insert_web_image',
            input: {
              ...input,
              _wiswork_image_base64: (scenario === 'normalized' ? jpegBytes : bytes).toString(
                'base64',
              ),
            },
          }),
        )
        if (scenario === 'normalized') {
          expect(prepareImageHandoff).toHaveBeenCalledWith({
            mime: 'image/png',
            bytes: Buffer.alloc(181 * 1024),
          })
          const forwarded = executeTool.mock.calls[0]![0] as any
          expect(Buffer.byteLength(JSON.stringify(forwarded.input))).toBeLessThan(256 * 1024)
        }
        expect(receipt.isError).toBe(false)
      } else if (scenario === 'remote-failed') {
        expect(executeTool).toHaveBeenCalledOnce()
        expect(receipt).toMatchObject({ isError: true, output: 'tool_execution_failed' })
      } else {
        expect(executeTool).not.toHaveBeenCalled()
        expect(receipt.isError).toBe(true)
        if (scenario === 'large' || scenario === 'normalizer-failed')
          expect(receipt.output).toBe('image_limit')
        if (scenario === 'unavailable') expect(receipt.output).toBe('image_fetch_unavailable')
        if (scenario !== 'cancelled') {
          const events = stream
            .split('\n')
            .filter((line) => line.startsWith('data: {'))
            .map((line) => JSON.parse(line.slice(6)))
          expect(events).toContainEqual(
            expect.objectContaining({
              type: 'wiswork_tool_lifecycle',
              tool_name: 'insert_web_image',
              state: 'error',
              summary: receipt.output,
            }),
          )
          expect(stream).not.toContain(input.url)
        }
        if (scenario === 'injected') {
          expect(rejectedBeforeApproval).toBe(true)
          expect(receipt.output).toBe('invalid_tool_input')
          expect(executeRetrieval).not.toHaveBeenCalled()
        }
      }
      expect(stream).not.toContain(bytes.toString('base64'))
      expect(stream).not.toContain(jpegBytes.toString('base64'))
      expect(stream).not.toContain('private upstream')
    },
  )

  it('keeps a remote Office mutation alive while the user reviews it for 68 seconds', async () => {
    vi.useFakeTimers()
    let session!: DocumentToolSession
    let receipt!: ToolExecution
    let resolveRemote!: (value: { output: string; isError: boolean }) => void
    const remote = new Promise<{ output: string; isError: boolean }>((resolve) => {
      resolveRemote = resolve
    })
    const proxy = createOfficeCodexProxy({
      runtime: {
        async runOfficeTurn(input: any) {
          session = input.toolSession
          const result = session.callTool(session.credentials, {
            id: 'slow-consent-write',
            name: 'edit_slide_text',
            input: {},
          })
          if (!isToolExecutionSuspension(result)) throw new Error('expected_mutation_suspension')
          receipt = await result.result
          input.onEvent({ type: 'terminal', status: 'completed' })
        },
      } as any,
      rollout,
      policyAuthority: createShellEnhancedPolicyAuthority(() => 0),
    })
    try {
      const response = await proxy({
        body: {
          system: '',
          messages: [],
          tools: [
            { name: 'edit_slide_text', description: 'write', input_schema: { type: 'object' } },
          ],
        },
        signal: new AbortController().signal,
        host: 'PowerPoint',
        sessionId: 'session_12345678',
        requestId: 'request_12345678',
        statement: { ...statement, host: 'office-powerpoint' },
        executeTool: vi.fn(() => remote),
      })
      await vi.advanceTimersByTimeAsync(68_000)
      resolveRemote({ output: 'applied', isError: false })
      for await (const _chunk of response.body as AsyncIterable<Uint8Array>) void _chunk
      expect(receipt).toMatchObject({ output: 'applied', isError: false })
    } finally {
      session?.close()
      vi.useRealTimers()
    }
  })

  it.each([
    ['cancel', 'success'],
    ['cancel', 'failure'],
    ['timeout', 'success'],
    ['timeout', 'failure'],
    ['active', 'failure'],
    ['active', 'invalid-result'],
  ] as const)(
    'contains a late mutation %s / %s without replaying its consumed claim',
    async (ending, outcome) => {
      vi.useFakeTimers()
      const unhandled = vi.fn()
      process.on('unhandledRejection', unhandled)
      let session!: DocumentToolSession
      let receipt!: ToolExecution
      let resolve!: (value: { output: string; isError: boolean }) => void
      let reject!: (error: Error) => void
      const remote = new Promise<{ output: string; isError: boolean }>((done, fail) => {
        resolve = done
        reject = fail
      })
      const finishRemote = () => {
        if (outcome === 'success') resolve({ output: 'late private receipt', isError: false })
        else if (outcome === 'invalid-result') resolve({ output: undefined, isError: false } as any)
        else reject(new Error('late private failure'))
      }
      const executeTool = vi.fn(() => remote)
      const proxy = createOfficeCodexProxy({
        runtime: {
          async runOfficeTurn(input: any) {
            session = input.toolSession
            const event = { callId: 'write_1', toolName: 'edit_slide_text' }
            input.onEvent({ type: 'tool-start', ...event })
            const result = session.callTool(session.credentials, {
              id: event.callId,
              name: event.toolName,
              input: {},
            })
            if (!isToolExecutionSuspension(result)) throw new Error('expected_mutation_suspension')
            receipt = await result.result
            input.onEvent({ type: 'tool-complete', ...event, isError: receipt.isError === true })
            input.onEvent({ type: 'terminal', status: 'completed' })
          },
        } as any,
        rollout,
        policyAuthority: createShellEnhancedPolicyAuthority(() => 0),
      })
      try {
        const response = await proxy({
          body: {
            system: '',
            messages: [],
            tools: [
              { name: 'edit_slide_text', description: 'write', input_schema: { type: 'object' } },
            ],
          },
          signal: new AbortController().signal,
          host: 'PowerPoint',
          sessionId: 'session_12345678',
          requestId: 'request_12345678',
          statement: { ...statement, host: 'office-powerpoint' },
          executeTool,
        })
        await vi.advanceTimersByTimeAsync(5)
        expect(executeTool).toHaveBeenCalledOnce()
        if (ending === 'cancel') session.cancelAll(session.credentials)
        else if (ending === 'timeout') await vi.advanceTimersByTimeAsync(5 * 60_000)
        else finishRemote()
        let stream = ''
        for await (const chunk of response.body as AsyncIterable<Uint8Array>)
          stream += new TextDecoder().decode(chunk)
        expect(receipt).toMatchObject({
          isError: true,
          output:
            ending === 'cancel'
              ? 'tool_cancelled'
              : ending === 'timeout'
                ? 'tool_timeout'
                : outcome === 'invalid-result'
                  ? 'invalid_tool_result'
                  : 'tool_execution_failed',
        })
        if (ending !== 'active') finishRemote()
        // Flush the detached pump chain, including the promise rejection notification turn.
        await vi.advanceTimersByTimeAsync(0)
        const events = stream
          .split('\n')
          .filter((line) => line.includes('wiswork_tool_lifecycle'))
          .map((line) => JSON.parse(line.slice(6)))
        expect(events.map((event) => event.state)).toEqual(['running', 'error'])
        expect(stream).not.toMatch(/late private|mutation_claim_consumed/)
        expect(unhandled).not.toHaveBeenCalled()
      } finally {
        process.removeListener('unhandledRejection', unhandled)
        session?.close()
        vi.useRealTimers()
      }
    },
  )

  it('closes an observed read on abort and ignores late semantic completion and private errors', async () => {
    const abort = new AbortController()
    let onEvent!: (event: any) => void
    let finish!: () => void
    const pending = new Promise<void>((resolve) => {
      finish = resolve
    })
    const proxy = createOfficeCodexProxy({
      runtime: {
        async runOfficeTurn(input: any) {
          onEvent = input.onEvent
          onEvent({
            type: 'tool-start',
            turnId: 'opaque_turn',
            callId: 'pending_read',
            toolName: 'get_document_text',
          })
          await pending
        },
      } as any,
      rollout,
      policyAuthority: createShellEnhancedPolicyAuthority(() => 0),
    })
    const response = await proxy({
      body: {
        system: '',
        messages: [],
        tools: [
          { name: 'get_document_text', description: 'read', input_schema: { type: 'object' } },
        ],
      },
      signal: abort.signal,
      host: 'Word',
      sessionId: 'session_12345678',
      requestId: 'request_12345678',
      statement,
      executeTool: vi.fn(),
    })
    abort.abort()
    onEvent({
      type: 'tool-complete',
      turnId: 'opaque_turn',
      callId: 'pending_read',
      toolName: 'get_document_text',
      isError: false,
      errorCode: 'private-secret',
    })
    onEvent({ type: 'terminal', status: 'completed' })
    onEvent({
      type: 'tool-start',
      turnId: 'opaque_turn',
      callId: 'late_read',
      toolName: 'get_document_text',
    })
    finish()
    let stream = ''
    for await (const chunk of response.body as AsyncIterable<Uint8Array>)
      stream += new TextDecoder().decode(chunk)
    const events = stream
      .split('\n')
      .filter((line) => line.includes('wiswork_tool_lifecycle'))
      .map((line) => JSON.parse(line.slice(6)))
    expect(events.map((event) => event.state)).toEqual(['running', 'error'])
    expect(stream).not.toContain('wiswork_tool_activity')
    expect(stream).not.toMatch(/private-secret|opaque_turn|pending_read|late_read/)
  })

  it('emits bounded SSE keepalives while a long Office turn is silent', async () => {
    vi.useFakeTimers()
    let finish!: () => void
    const pending = new Promise<void>((resolve) => {
      finish = resolve
    })
    const proxy = createOfficeCodexProxy({
      runtime: { runOfficeTurn: vi.fn(() => pending) } as any,
      rollout,
      policyAuthority: createShellEnhancedPolicyAuthority(() => 0),
    })
    const response = await proxy({
      body: {
        system: 'rules',
        messages: [],
        tools: [
          { name: 'get_document_text', description: 'read', input_schema: { type: 'object' } },
        ],
      },
      signal: new AbortController().signal,
      host: 'Word',
      sessionId: 'session_12345678',
      requestId: 'request_12345678',
      statement,
      executeTool: vi.fn(),
    })
    const iterator = (response.body as AsyncIterable<Uint8Array>)[Symbol.asyncIterator]()
    const next = iterator.next()
    await vi.advanceTimersByTimeAsync(OFFICE_PROXY_KEEPALIVE_MS)
    await expect(next).resolves.toMatchObject({
      done: false,
      value: expect.any(Uint8Array),
    })
    finish()
    await iterator.return?.()
    vi.useRealTimers()
  })

  it('does not report success when a failed terminal event precedes promise rejection', async () => {
    let finish!: () => void
    const pending = new Promise<void>((resolve) => {
      finish = resolve
    })
    const proxy = createOfficeCodexProxy({
      runtime: {
        async runOfficeTurn(input: any) {
          input.onEvent({ type: 'terminal', status: 'failed' })
          await pending
          throw new Error('enhanced_request_rejected')
        },
      } as any,
      rollout,
      policyAuthority: createShellEnhancedPolicyAuthority(() => 0),
    })
    const response = await proxy({
      body: {
        system: 'rules',
        messages: [{ role: 'user', content: 'read' }],
        tools: [
          { name: 'get_document_text', description: 'read', input_schema: { type: 'object' } },
        ],
      },
      signal: new AbortController().signal,
      host: 'Word',
      sessionId: 'session_12345678',
      requestId: 'request_12345678',
      statement,
      executeTool: vi.fn(),
    })
    const chunks: string[] = []
    const consume = (async () => {
      for await (const chunk of response.body as AsyncIterable<Uint8Array>)
        chunks.push(new TextDecoder().decode(chunk))
    })()
    const rejected = expect(consume).rejects.toThrow('enhanced_request_rejected')
    await Promise.resolve()
    finish()
    await rejected
    expect(chunks.join('')).not.toContain('[DONE]')
  })
  it('uses one document-scoped tool session and returns Codex text as bounded SSE', async () => {
    const executeTool = vi.fn(async () => ({ output: '{"text":"hello"}', isError: false }))
    const runtime = {
      async runOfficeTurn(input: any) {
        const result = await input.toolSession.callTool(input.toolSession.credentials, {
          id: 'call_12345678',
          name: 'get_document_text',
          input: {},
        })
        expect(result.output).toBe('{"text":"hello"}')
        input.onEvent({ type: 'text', text: 'Done' })
        input.onEvent({ type: 'terminal', status: 'completed' })
      },
    }
    const authority = createShellEnhancedPolicyAuthority(() => 0)
    const telemetry = { component: vi.fn(), host: vi.fn() }
    const proxy = createOfficeCodexProxy({
      runtime: runtime as any,
      rollout,
      policyAuthority: authority,
      telemetry,
    })
    const response = await proxy({
      body: {
        system: 'Office rules',
        messages: [{ role: 'user', content: 'read' }],
        tools: [
          { name: 'get_document_text', description: 'read', input_schema: { type: 'object' } },
        ],
      },
      signal: new AbortController().signal,
      host: 'Word',
      sessionId: 'session_12345678',
      requestId: 'request_12345678',
      statement,
      executeTool,
    })
    const chunks: string[] = []
    for await (const chunk of response.body as AsyncIterable<Uint8Array>)
      chunks.push(new TextDecoder().decode(chunk))
    expect(chunks.join('')).toContain('Done')
    expect(chunks.join('')).not.toContain('wiswork_tool_activity')
    expect(executeTool).toHaveBeenCalledWith(
      expect.objectContaining({ generation: 3, toolName: 'get_document_text' }),
    )
    expect(telemetry.host.mock.calls).toEqual([
      ['office-word', 'plan', 'started'],
      ['office-word', 'plan', 'succeeded'],
      ['office-word', 'dispatch', 'started'],
      ['office-word', 'dispatch', 'succeeded'],
      ['office-word', 'verify', 'verified'],
      ['office-word', 'complete', 'succeeded'],
    ])
  })

  it('filters unknown/shared tools and raw Office independently from the semantic allowlist', async () => {
    const runtime = {
      async runOfficeTurn(input: any) {
        expect(
          input.toolSession.listTools(input.toolSession.credentials).map((tool: any) => tool.name),
        ).toEqual(['get_document_text', 'execute_office_js'])
        input.onEvent({ type: 'terminal', status: 'completed' })
      },
    }
    const proxy = createOfficeCodexProxy({
      runtime: runtime as any,
      rollout,
      policyAuthority: createShellEnhancedPolicyAuthority(() => 0),
    })
    const response = await proxy({
      body: {
        system: 'rules',
        messages: [],
        tools: [
          { name: 'get_document_text', description: 'read', input_schema: { type: 'object' } },
          {
            name: 'execute_office_js',
            description: 'bounded declarative operations',
            input_schema: { type: 'object' },
          },
          { name: 'bash', description: 'forbidden', input_schema: { type: 'object' } },
        ],
      },
      signal: new AbortController().signal,
      host: 'Word',
      sessionId: 'session_12345678',
      requestId: 'request_12345678',
      statement,
      executeTool: vi.fn(),
    })
    for await (const _chunk of response.body as AsyncIterable<Uint8Array>) {
      /* drain */
    }
  })

  it.each(['Mac', 'PC'])(
    'preserves the real PowerPoint %s tool inventory and dispatches screenshot reviews',
    async (platform) => {
      const skill = createPowerPointSkill({
        adapter: {} as PowerPointAdapter,
        proposals: createStructuredProposalController(),
        platform,
        nativeMasterEditingSupported: true,
      })
      let registered: ReturnType<DocumentToolSession['listTools']> = []
      let review: ToolExecution | undefined
      const reviewInput = { slide_index: 0, acceptance_ids: ['A1.1'], passed: true }
      const executeTool = vi.fn(async () => ({
        output: 'design_contract_screenshot_required',
        isError: true,
      }))
      const proxy = createOfficeCodexProxy({
        runtime: {
          async runOfficeTurn(input: any) {
            const session: DocumentToolSession = input.toolSession
            registered = session.listTools(session.credentials)
            const result = session.callTool(session.credentials, {
              id: 'review01',
              name: 'review_slide_screenshot',
              input: reviewInput,
            })
            review = await (isToolExecutionSuspension(result) ? result.result : result)
            input.onEvent({ type: 'terminal', status: 'completed' })
          },
        } as any,
        rollout,
        policyAuthority: createShellEnhancedPolicyAuthority(() => 0),
      })
      const response = await proxy({
        // Match the actual request wire: shared schema references become detached JSON values.
        body: JSON.parse(
          JSON.stringify({
            system: skill.systemPrompt,
            messages: [],
            tools: skill.tools.map((tool) => ({
              name: tool.name,
              description: tool.description,
              input_schema: tool.inputSchema,
            })),
          }),
        ),
        signal: new AbortController().signal,
        host: 'PowerPoint',
        sessionId: 'session_12345678',
        requestId: 'request_12345678',
        statement: { ...statement, host: 'office-powerpoint' },
        executeTool,
      })
      for await (const _chunk of response.body as AsyncIterable<Uint8Array>) {
        /* drain */
      }

      expect(registered.map((tool) => tool.name)).toEqual(skill.tools.map((tool) => tool.name))
      expect(registered.find((tool) => tool.name === 'review_slide_screenshot')).toMatchObject({
        ...skill.tools.find((tool) => tool.name === 'review_slide_screenshot'),
        annotations: { readOnlyHint: true, destructiveHint: false },
      })
      expect(executeTool).toHaveBeenCalledExactlyOnceWith(
        expect.objectContaining({ toolName: 'review_slide_screenshot', input: reviewInput }),
      )
      expect(review).toMatchObject({
        output: 'design_contract_screenshot_required',
        isError: true,
        mutated: false,
      })
    },
  )

  it('does not claim a legacy metadata-only screenshot was delivered to the model', async () => {
    let screenshot: ToolExecution | undefined
    const proxy = createOfficeCodexProxy({
      runtime: {
        async runOfficeTurn(input: any) {
          screenshot = await input.toolSession.callTool(input.toolSession.credentials, {
            id: 'screenshot01',
            name: 'screenshot_slide',
            input: { slide_index: 0 },
          })
          input.onEvent({ type: 'terminal', status: 'completed' })
        },
      } as any,
      rollout,
      policyAuthority: createShellEnhancedPolicyAuthority(() => 0),
    })
    const response = await proxy({
      body: {
        system: '',
        messages: [],
        tools: [{ name: 'screenshot_slide', description: '', input_schema: { type: 'object' } }],
      },
      signal: new AbortController().signal,
      host: 'PowerPoint',
      sessionId: 'session_12345678',
      requestId: 'request_12345678',
      statement: { ...statement, host: 'office-powerpoint' },
      executeTool: async () => ({
        output: JSON.stringify({ mime: 'image/png', bytes: 100, visualAvailableToModel: true }),
        isError: false,
      }),
    })
    for await (const _chunk of response.body as AsyncIterable<Uint8Array>) {
      /* drain */
    }
    expect(screenshot).toMatchObject({
      output: 'office_screenshot_unavailable',
      isError: true,
      mutated: false,
    })
    expect(screenshot?.modelContent).toBeUndefined()
  })

  it.each(['unavailable', 'decoder failed', 'changed bytes', 'changed MIME', 'cancelled'])(
    'refuses to promote a screenshot when native validation is %s',
    async (scenario) => {
      const image = {
        mime: 'image/png',
        base64:
          'iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAYAAAAfFcSJAAAAC0lEQVR4AWP4DwQACfsD/c8LaHIAAAAASUVORK5CYII=',
      }
      const controller = new AbortController()
      let screenshot: ToolExecution | undefined
      const proxy = createOfficeCodexProxy({
        rollout,
        policyAuthority: createShellEnhancedPolicyAuthority(() => 0),
        ...(scenario === 'unavailable'
          ? {}
          : {
              prepareImageHandoff: async (value: {
                bytes: Uint8Array
                mime: 'image/png' | 'image/jpeg'
              }) => {
                if (scenario === 'decoder failed') throw new Error('private native decoder error')
                if (scenario === 'cancelled') controller.abort()
                return {
                  bytes: scenario === 'changed bytes' ? new Uint8Array() : value.bytes,
                  mime: scenario === 'changed MIME' ? ('image/jpeg' as const) : value.mime,
                }
              },
            }),
        runtime: {
          async runOfficeTurn(input: any) {
            screenshot = await input.toolSession.callTool(input.toolSession.credentials, {
              id: 'screenshot01',
              name: 'screenshot_slide',
              input: { slide_index: 0 },
            })
            input.onEvent({ type: 'terminal', status: 'completed' })
          },
        } as any,
      })
      const response = await proxy({
        body: {
          system: '',
          messages: [],
          tools: [{ name: 'screenshot_slide', description: '', input_schema: { type: 'object' } }],
        },
        signal: controller.signal,
        host: 'PowerPoint',
        sessionId: 'session_12345678',
        requestId: 'request_12345678',
        statement: { ...statement, host: 'office-powerpoint' },
        executeTool: async () => ({
          output: encodeOfficeScreenshotResult('{}', [{ type: 'image', image }]),
          isError: false,
        }),
      })
      for await (const _chunk of response.body as AsyncIterable<Uint8Array>) {
        /* drain */
      }
      expect(screenshot).toMatchObject({
        isError: true,
        mutated: false,
        output: scenario === 'cancelled' ? 'tool_cancelled' : 'office_screenshot_unavailable',
      })
      expect(screenshot?.modelContent).toBeUndefined()
    },
  )

  it('executes the Enhanced PowerPoint state, feedback, and planning sequence', async () => {
    const executeTool = vi.fn(async (call: { toolName: string; callId: string }) => {
      if (!/^[A-Za-z0-9_-]{8,128}$/.test(call.callId)) throw new Error('invalid_tool_call')
      return {
        output:
          call.toolName === 'get_presentation_state'
            ? '{"slideCount":1,"selectedSlideIndexes":[]}'
            : call.toolName === 'ask_clarification'
              ? 'audience: general; style: concise; pages: 8'
              : '{"title":"LLM","slides":[]}',
        isError: false,
      }
    })
    const runtime = {
      async runOfficeTurn(input: any) {
        expect(
          input.toolSession.listTools(input.toolSession.credentials).map((tool: any) => tool.name),
        ).toEqual(['get_presentation_state', 'ask_clarification', 'plan_deck', 'verify_slides'])
        for (const name of ['get_presentation_state', 'ask_clarification', 'plan_deck']) {
          const result = await input.toolSession.callTool(input.toolSession.credentials, {
            id:
              name === 'get_presentation_state'
                ? 'state01'
                : name === 'ask_clarification'
                  ? 'aud01'
                  : 'p',
            name,
            input: {},
          })
          expect(result.isError).not.toBe(true)
        }
        input.onEvent({ type: 'terminal', status: 'completed' })
      },
    }
    const proxy = createOfficeCodexProxy({
      runtime: runtime as any,
      rollout,
      policyAuthority: createShellEnhancedPolicyAuthority(() => 0),
    })
    const response = await proxy({
      body: {
        system: 'PowerPoint rules',
        messages: [{ role: 'user', content: 'Create a six-slide deck' }],
        tools: [
          {
            name: 'get_presentation_state',
            description: 'state',
            input_schema: { type: 'object' },
          },
          { name: 'ask_clarification', description: 'feedback', input_schema: { type: 'object' } },
          { name: 'plan_deck', description: 'plan', input_schema: { type: 'object' } },
          { name: 'verify_slides', description: 'verify', input_schema: { type: 'object' } },
          {
            name: 'get_document_text',
            description: 'wrong host',
            input_schema: { type: 'object' },
          },
        ],
      },
      signal: new AbortController().signal,
      host: 'PowerPoint',
      sessionId: 'session_12345678',
      requestId: 'request_12345678',
      statement: { ...statement, host: 'office-powerpoint' },
      executeTool,
    })
    for await (const _chunk of response.body as AsyncIterable<Uint8Array>) {
      /* drain */
    }
    expect(executeTool.mock.calls.map(([call]) => call.toolName)).toEqual([
      'get_presentation_state',
      'ask_clarification',
      'plan_deck',
    ])
    expect(new Set(executeTool.mock.calls.map(([call]) => call.callId)).size).toBe(3)
  })

  it('keeps PC-backed web and image search available to Enhanced PowerPoint turns', async () => {
    const executeTool = vi.fn()
    const executeRetrieval = vi.fn(async (capability: string) =>
      new TextEncoder().encode(
        JSON.stringify(
          capability === 'image-search.v1'
            ? {
                images: [
                  {
                    title: 'Volcano',
                    source_url: 'https://example.com/volcano',
                    image_url: 'https://example.com/image.jpg',
                    private_field: 'private-payload',
                  },
                ],
              }
            : capability === 'web-fetch.v1'
              ? {
                  url: 'https://example.com/article',
                  title: 'Volcano article',
                  content: 'private-page-content',
                  content_type: 'text/plain',
                }
              : { results: [] },
        ),
      ),
    )
    const runtime = {
      async runOfficeTurn(input: any) {
        const names = input.toolSession
          .listTools(input.toolSession.credentials)
          .map((tool: any) => tool.name)
        expect(names).toEqual(['web_search', 'web_fetch', 'image_search', 'plan_deck'])
        for (const name of names.slice(0, 3)) {
          input.onEvent({ type: 'text', text: `before ${name}` })
          const result = await semanticCall(input, {
            id: `call_${name}`,
            name,
            input:
              name === 'web_fetch'
                ? { url: 'https://example.com' }
                : { query: 'volcano', max_results: 4 },
          })
          expect(result.isError).toBe(false)
          if (name === 'image_search') expect(result.output).toContain('private-payload')
          input.onEvent({ type: 'text', text: `after ${name}` })
        }
        input.onEvent({ type: 'terminal', status: 'completed' })
      },
    }
    const proxy = createOfficeCodexProxy({
      runtime: runtime as any,
      rollout,
      policyAuthority: createShellEnhancedPolicyAuthority(() => 0),
    })
    const response = await proxy({
      body: {
        system: 'PowerPoint rules',
        messages: [{ role: 'user', content: 'Create a researched deck' }],
        tools: ['web_search', 'web_fetch', 'image_search', 'plan_deck'].map((name) => ({
          name,
          description: name,
          input_schema: { type: 'object' },
        })),
      },
      signal: new AbortController().signal,
      host: 'PowerPoint',
      sessionId: 'session_12345678',
      requestId: 'request_12345678',
      statement: { ...statement, host: 'office-powerpoint' },
      executeRetrieval,
      executeTool,
    })
    let stream = ''
    for await (const chunk of response.body as AsyncIterable<Uint8Array>)
      stream += new TextDecoder().decode(chunk)
    const activities = stream
      .split('\n')
      .filter((line) => line.startsWith('data: {'))
      .map((line) => JSON.parse(line.slice(6)))
      .filter((event) => event.type === 'wiswork_tool_activity')
    expect(activities.map((event) => [event.tool_name, event.state])).toEqual([
      ['web_search', 'running'],
      ['web_search', 'complete'],
      ['web_fetch', 'running'],
      ['web_fetch', 'complete'],
      ['image_search', 'running'],
      ['image_search', 'complete'],
    ])
    expect(activities[5]).toMatchObject({
      generation: 3,
      query: 'volcano',
      started_at: expect.any(Number),
    })
    expect(activities[3]).toMatchObject({
      tool_name: 'web_fetch',
      state: 'complete',
      display: {
        kind: 'links',
        items: [{ url: 'https://example.com/article', title: 'Volcano article' }],
      },
    })
    expect(activities[3]).not.toHaveProperty('result_count')
    expect(stream).not.toContain('private-page-content')
    expect(activities[5]).toMatchObject({
      call_id: activities[4].call_id,
      result_count: 1,
      display: {
        kind: 'images',
        items: [{ title: 'Volcano', url: 'https://example.com/volcano' }],
      },
    })
    expect(stream).not.toContain('private-payload')
    expect(stream).not.toContain('image.jpg')
    expect(stream).not.toContain('tool_use')
    expect(stream.indexOf('before image_search')).toBeLessThan(
      stream.indexOf('"tool_name":"image_search","state":"running"'),
    )
    expect(stream.indexOf('after image_search')).toBeGreaterThan(
      stream.indexOf('"tool_name":"image_search","state":"complete"'),
    )
    expect(executeTool).not.toHaveBeenCalled()
    expect(executeRetrieval.mock.calls.map(([capability]) => capability)).toEqual([
      'web-search.v1',
      'web-fetch.v1',
      'image-search.v1',
    ])
  })

  it('returns a recoverable tool error when PC-backed image search is unavailable', async () => {
    const runtime = {
      async runOfficeTurn(input: any) {
        const result = await semanticCall(input, {
          id: 'call_image_search',
          name: 'image_search',
          input: { query: 'volcano', max_results: 4 },
        })
        expect(result).toMatchObject({ isError: true, output: 'retrieval_upstream_error' })
        input.onEvent({ type: 'terminal', status: 'completed' })
      },
    }
    const proxy = createOfficeCodexProxy({
      runtime: runtime as any,
      rollout,
      policyAuthority: createShellEnhancedPolicyAuthority(() => 0),
    })
    const response = await proxy({
      body: {
        system: 'PowerPoint rules',
        messages: [],
        tools: [{ name: 'image_search', description: 'images', input_schema: { type: 'object' } }],
      },
      signal: new AbortController().signal,
      host: 'PowerPoint',
      sessionId: 'session_12345678',
      requestId: 'request_12345678',
      statement: { ...statement, host: 'office-powerpoint' },
      executeRetrieval: vi.fn(async () => {
        throw new Error('private-upstream-secret')
      }),
      executeTool: vi.fn(),
    })
    let stream = ''
    for await (const chunk of response.body as AsyncIterable<Uint8Array>)
      stream += new TextDecoder().decode(chunk)
    const activities = stream
      .split('\n')
      .filter((line) => line.startsWith('data: {'))
      .map((line) => JSON.parse(line.slice(6)))
      .filter((event) => event.type === 'wiswork_tool_activity')
    expect(activities.map((event) => event.state)).toEqual(['running', 'error'])
    expect(activities[1]).toMatchObject({
      call_id: activities[0].call_id,
      summary: 'Retrieval unavailable',
    })
    expect(stream).not.toContain('private-upstream-secret')
  })

  it('bounds retrieval cards and excludes private fields, credentials and unsafe source URLs', async () => {
    const images = [
      { title: 'private', source_url: 'https://user:secret@example.com/' },
      { title: 'local', source_url: 'https://127.0.0.1/' },
      { title: 'script', source_url: 'javascript:alert(1)' },
      ...Array.from({ length: 20 }, (_, index) => ({
        title: '图'.repeat(500),
        source_url: `https://example.com/${index}/${'x'.repeat(1500)}?token=private-token#secret`,
        image_url: 'https://example.com/private-preview',
        private_body: 'private-payload',
      })),
    ]
    const output = JSON.stringify({ images, authorization: 'private-auth' })
    const executeTool = vi.fn()
    const proxy = createOfficeCodexProxy({
      runtime: {
        async runOfficeTurn(input: any) {
          const result = await semanticCall(input, {
            id: 'images',
            name: 'image_search',
            input: { query: '图'.repeat(4096), max_results: 20 },
          })
          expect(result).toMatchObject({ isError: false, output })
          input.onEvent({ type: 'terminal', status: 'completed' })
        },
      } as any,
      rollout,
      policyAuthority: createShellEnhancedPolicyAuthority(() => 0),
    })
    const response = await proxy({
      body: {
        system: '',
        messages: [],
        tools: [{ name: 'image_search', description: 'images', input_schema: { type: 'object' } }],
      },
      signal: new AbortController().signal,
      host: 'PowerPoint',
      sessionId: 'session_12345678',
      requestId: 'request_12345678',
      statement: { ...statement, host: 'office-powerpoint' },
      executeTool,
      executeRetrieval: async () => new TextEncoder().encode(output),
    })
    let stream = ''
    for await (const chunk of response.body as AsyncIterable<Uint8Array>)
      stream += new TextDecoder().decode(chunk)
    const lines = stream.split('\n').filter((line) => line.includes('wiswork_tool_activity'))
    const [start, complete] = lines.map((line) => JSON.parse(line.slice(6)))
    expect(start).not.toHaveProperty('query')
    expect(complete.query).toHaveLength(240)
    expect(complete.result_count).toBe(20)
    expect(complete.display.items.length).toBeGreaterThan(0)
    expect(complete.display.items.length).toBeLessThanOrEqual(8)
    for (const line of lines) expect(Buffer.byteLength(line)).toBeLessThan(12 * 1024)
    expect(stream).not.toMatch(
      /private-payload|private-token|private-auth|private-preview|127\.0\.0\.1|javascript:|user:secret/,
    )
    expect(executeTool).not.toHaveBeenCalled()
  })

  it('keeps bounded declarative PowerPoint writes without granting raw Office authority', async () => {
    const runtime = {
      async runOfficeTurn(input: any) {
        expect(
          input.toolSession.listTools(input.toolSession.credentials).map((tool: any) => tool.name),
        ).toEqual(['execute_office_js', 'edit_slide_text'])
        expect(
          input.summarizeProposal({
            id: 'proposal-1',
            name: 'execute_office_js',
            input: { program: { version: 1, operations: [{ op: 'add_text_box' }] } },
          }),
        ).toEqual({ operation: 'replace', target: 'slides', scope: 'bounded-set', count: 1 })
        expect(
          input.summarizeProposal({
            id: 'proposal-2',
            name: 'edit_slide_text',
            input: { edits: [{ slide_index: 0, shape_id: 'shape-1', text: 'Hello' }] },
          }),
        ).toEqual({ operation: 'replace', target: 'slides', scope: 'bounded-set', count: 1 })
        input.onEvent({ type: 'terminal', status: 'completed' })
      },
    }
    const proxy = createOfficeCodexProxy({
      runtime: runtime as any,
      rollout,
      policyAuthority: createShellEnhancedPolicyAuthority(() => 0),
    })
    const response = await proxy({
      body: {
        system: 'PowerPoint rules',
        messages: [{ role: 'user', content: 'Create a presentation' }],
        tools: [
          {
            name: 'execute_office_js',
            description: 'bounded declarative operations',
            input_schema: { type: 'object' },
          },
          {
            name: 'edit_slide_text',
            description: 'bounded text edit',
            input_schema: { type: 'object' },
          },
          {
            name: 'propose_raw_office_edit',
            description: 'elevated raw edit',
            input_schema: { type: 'object' },
          },
        ],
      },
      signal: new AbortController().signal,
      host: 'PowerPoint',
      sessionId: 'session_12345678',
      requestId: 'request_12345678',
      statement: { ...statement, host: 'office-powerpoint', raw_office: false },
      executeTool: vi.fn(),
    })
    for await (const _chunk of response.body as AsyncIterable<Uint8Array>) {
      /* drain */
    }
  })

  it('exposes only the distinct raw proposal tool when both signed statement and trusted policy allow it', async () => {
    const runtime = {
      async runOfficeTurn(input: any) {
        expect(
          input.toolSession.listTools(input.toolSession.credentials).map((tool: any) => tool.name),
        ).toEqual(['propose_raw_office_edit'])
        input.onEvent({ type: 'terminal', status: 'completed' })
      },
    }
    const proxy = createOfficeCodexProxy({
      runtime: runtime as any,
      rollout: { ...rollout, rawOfficeEnabled: true },
      policyAuthority: createShellEnhancedPolicyAuthority(() => 0),
    })
    const response = await proxy({
      body: {
        system: 'rules',
        messages: [],
        tools: [
          { name: 'propose_raw_office_edit', description: 'raw', input_schema: { type: 'object' } },
        ],
      },
      signal: new AbortController().signal,
      host: 'Word',
      sessionId: 'session_12345678',
      requestId: 'request_12345678',
      statement: { ...statement, raw_office: true },
      executeTool: vi.fn(),
    })
    for await (const _chunk of response.body as AsyncIterable<Uint8Array>) {
      /* drain */
    }
  })
})
