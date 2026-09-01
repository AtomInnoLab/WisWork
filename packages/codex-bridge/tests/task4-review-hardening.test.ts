import { isToolExecutionSuspension, suspendToolExecution } from '@wiswork/agent-core'
import type { EnhancedCapability } from '@wiswork/agent-runtime'
import { describe, expect, it, vi } from 'vitest'
import {
  createDocumentToolManifest,
  createDocumentToolSession,
  type DocumentToolRegistration,
} from '../src/tool-router.js'

const policy = {
  globalEnabled: true,
  rawOfficeEnabled: false,
  hosts: {
    latex: true,
    slides: true,
    docs: true,
    sheets: true,
    'office-word': true,
    'office-excel': true,
    'office-powerpoint': true,
  },
}
const read = { name: 'get_document_context', description: 'Read.', inputSchema: { type: 'object' } }
const mutate = { name: 'replace_blocks', description: 'Replace.', inputSchema: { type: 'object' } }
function policyGrant(
  capabilities: EnhancedCapability[] = ['semantic-read', 'transaction-proposal'],
) {
  const grant = Object.freeze({})
  const snapshot = { generation: 1, host: 'docs' as const, policy, capabilities }
  return {
    policyGrant: grant,
    consumePolicyGrant: (candidate: unknown) => {
      if (candidate !== grant) throw new Error('invalid_enhanced_policy_handle')
      return snapshot
    },
  }
}

function registration(overrides: Partial<DocumentToolRegistration> = {}): DocumentToolRegistration {
  const manifest = createDocumentToolManifest({
    ...policyGrant(),
    tools: [read, mutate],
    policy: { get_document_context: 'read', replace_blocks: 'mutate' },
  })
  return {
    identity: {
      ownerId: 'owner',
      host: 'docs',
      documentId: 'doc',
      sessionId: 'session',
      generation: 1,
    },
    manifest,
    isOpen: () => true,
    executeRead: vi.fn(async () => ({ output: 'text', summary: 'read' })),
    suspendMutation: suspendToolExecution,
    ...overrides,
  }
}

describe('Task 4 review hardening', () => {
  it('rejects aliases and capabilities not compiled for the exact host', () => {
    expect(() =>
      createDocumentToolManifest({
        ...policyGrant(['semantic-read']),
        tools: [{ ...read, name: 'read_document' }],
        policy: { read_document: 'read' },
      }),
    ).toThrow('tool_not_compiled_for_host')
    expect(() =>
      createDocumentToolManifest({
        ...policyGrant(['semantic-read']),
        tools: [mutate],
        policy: { replace_blocks: 'mutate' },
      }),
    ).toThrow('tool_capability_denied')
  })

  it('accepts only strict Shell-authorized policy snapshots', () => {
    const input = {
      ...policyGrant(['semantic-read']),
      tools: [read],
      policy: { get_document_context: 'read' as const },
    }
    expect(() => createDocumentToolManifest(input)).not.toThrow()
    expect(() =>
      createDocumentToolManifest({
        ...input,
        policyGrant: {} as any,
      }),
    ).toThrow('invalid_enhanced_policy_handle')
  })

  it('has no mutation callback injection point and requires opaque claim before settle', async () => {
    expect(() =>
      createDocumentToolSession({
        ...registration(),
        prepareMutation: () => {
          throw new Error('write')
        },
      } as any),
    ).toThrow('invalid_tool_session')
    const first = createDocumentToolSession(registration())
    const second = createDocumentToolSession(registration())
    const outcome = first.callTool(first.credentials, {
      id: 'w',
      name: 'replace_blocks',
      input: { text: 'new' },
    })
    expect(isToolExecutionSuspension(outcome as any)).toBe(true)
    const claimed = first.mutationAuthority.claimNext()!
    expect(claimed.request.call.input).toEqual({ text: 'new' })
    expect(() =>
      second.mutationAuthority.settle(claimed.claim, {
        output: 'bad',
        summary: 'bad',
        mutated: true,
      }),
    ).toThrow('mutation_claim_issuer_mismatch')
    expect(() =>
      first.mutationAuthority.settle({} as never, { output: 'bad', summary: 'bad' }),
    ).toThrow('invalid_mutation_claim')
    first.mutationAuthority.settle(claimed.claim, {
      output: 'changed',
      summary: 'changed',
      mutated: true,
    })
    expect(() =>
      first.mutationAuthority.settle(claimed.claim, {
        output: 'again',
        summary: 'again',
        mutated: true,
      }),
    ).toThrow('mutation_claim_consumed')
    await expect((outcome as any).result).resolves.toMatchObject({ output: 'changed' })
  })

  it('permanently consumes call ids and closes after the bounded total-call budget', async () => {
    const session = createDocumentToolSession(registration({ maxTotalCalls: 2 }))
    const call = { id: 'same', name: 'get_document_context', input: {} }
    await session.callTool(session.credentials, call)
    expect(session.callTool(session.credentials, call)).toMatchObject({
      output: 'tool_call_consumed',
      isError: true,
    })
    await session.callTool(session.credentials, { ...call, id: 'second' })
    expect(session.callTool(session.credentials, { ...call, id: 'third' })).toMatchObject({
      output: 'tool_session_call_limit',
      isError: true,
    })
    expect(() => session.listTools(session.credentials)).toThrow('tool_session_closed')
  })

  it('bounds the detached mutation queue and cancels unresolved work on close', async () => {
    const session = createDocumentToolSession(registration({ maxPendingMutations: 1 }))
    const first = session.callTool(session.credentials, {
      id: 'first',
      name: 'replace_blocks',
      input: {},
    })
    expect(isToolExecutionSuspension(first as any)).toBe(true)
    expect(
      session.callTool(session.credentials, { id: 'second', name: 'replace_blocks', input: {} }),
    ).toMatchObject({ output: 'mutation_queue_full', isError: true })
    session.close()
    await expect((first as any).result).resolves.toMatchObject({
      output: 'tool_cancelled',
      isError: true,
    })
    expect(session.mutationAuthority.claimNext()).toBeUndefined()
  })

  it('rejects aggregate catalog amplification before MCP serialization', () => {
    const names = [
      'get_document_context',
      'read_blocks',
      'insert_content',
      'replace_blocks',
      'apply_commands',
      'insert_image',
      'insert_chart',
      'edit_chart',
    ]
    const tools = names.map((name) => ({
      ...read,
      name,
      description: 'x'.repeat(100_000),
    }))
    expect(() =>
      createDocumentToolManifest({
        ...policyGrant(),
        tools,
        policy: Object.fromEntries(
          tools.map((tool) => [
            tool.name,
            ['get_document_context', 'read_blocks'].includes(tool.name) ? 'read' : 'mutate',
          ]),
        ),
      }),
    ).toThrow('tool_catalog_limit')
  })
})
