import { describe, expect, it, vi } from 'vitest'
import { ENHANCED_HOSTS, type EnhancedRolloutPolicy } from '@wiswork/agent-runtime'
import { OFFICE_PROXY_KEEPALIVE_MS, createOfficeCodexProxy } from '../src/main/office-codex-proxy'
import { createShellEnhancedPolicyAuthority } from '../src/main/enhanced-policy-authority'

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
      capability === 'image-search.v1' ? { images: [] } : { results: [] },
    )
    const runtime = {
      async runOfficeTurn(input: any) {
        const names = input.toolSession
          .listTools(input.toolSession.credentials)
          .map((tool: any) => tool.name)
        expect(names).toEqual(['web_search', 'web_fetch', 'image_search', 'plan_deck'])
        for (const name of names.slice(0, 3)) {
          await input.toolSession.callTool(input.toolSession.credentials, {
            id: `call_${name}`,
            name,
            input:
              name === 'web_fetch'
                ? { url: 'https://example.com' }
                : { query: 'volcano', max_results: 4 },
          })
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
    for await (const _chunk of response.body as AsyncIterable<Uint8Array>) {
      /* drain */
    }
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
        const result = await input.toolSession.callTool(input.toolSession.credentials, {
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
        throw new Error('retrieval_upstream_error')
      }),
      executeTool: vi.fn(),
    })
    for await (const _chunk of response.body as AsyncIterable<Uint8Array>) {
      /* drain */
    }
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
