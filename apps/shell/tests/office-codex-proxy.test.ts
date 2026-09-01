import { describe, expect, it, vi } from 'vitest'
import { ENHANCED_HOSTS, type EnhancedRolloutPolicy } from '@wiswork/agent-runtime'
import { createOfficeCodexProxy } from '../src/main/office-codex-proxy'
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
        ).toEqual(['get_document_text'])
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
          { name: 'execute_office_js', description: 'raw', input_schema: { type: 'object' } },
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
