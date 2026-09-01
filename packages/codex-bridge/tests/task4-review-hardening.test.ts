import { suspendToolExecution } from '@wiswork/agent-core'
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

function registration(overrides: Partial<DocumentToolRegistration> = {}): DocumentToolRegistration {
  const manifest = createDocumentToolManifest({
    authorization: {
      host: 'docs',
      policy,
      declaration: { capabilities: ['semantic-read', 'transaction-proposal'] },
    },
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
    prepareMutation: vi.fn(() =>
      suspendToolExecution(
        Promise.resolve({ output: 'changed', summary: 'changed', mutated: true }),
      ),
    ),
    ...overrides,
  }
}

describe('Task 4 review hardening', () => {
  it('rejects aliases and capabilities not compiled for the exact host', () => {
    expect(() =>
      createDocumentToolManifest({
        authorization: { host: 'docs', policy, declaration: { capabilities: ['semantic-read'] } },
        tools: [{ ...read, name: 'read_document' }],
        policy: { read_document: 'read' },
      }),
    ).toThrow('tool_not_compiled_for_host')
    expect(() =>
      createDocumentToolManifest({
        authorization: { host: 'docs', policy, declaration: { capabilities: ['semantic-read'] } },
        tools: [mutate],
        policy: { replace_blocks: 'mutate' },
      }),
    ).toThrow('tool_capability_denied')
  })

  it('does not accept a shape-forged mutation suspension', async () => {
    const session = createDocumentToolSession(
      registration({
        prepareMutation: vi.fn(
          () =>
            ({
              kind: 'tool-execution-suspension',
              result: Promise.resolve({ output: 'forged', summary: 'forged', mutated: true }),
              output: 'tool_execution_suspended',
              summary: 'Awaiting tool execution',
            }) as any,
        ),
      }),
    )
    await expect(
      session.callTool(session.credentials, { id: 'w', name: 'replace_blocks', input: {} }),
    ).resolves.toMatchObject({ output: 'mutation_authority_required', isError: true })
  })

  it('permanently consumes call ids and closes after the bounded total-call budget', async () => {
    const session = createDocumentToolSession(registration({ maxTotalCalls: 2 }))
    const call = { id: 'same', name: 'get_document_context', input: {} }
    await session.callTool(session.credentials, call)
    await expect(session.callTool(session.credentials, call)).resolves.toMatchObject({
      output: 'tool_call_consumed',
      isError: true,
    })
    await session.callTool(session.credentials, { ...call, id: 'second' })
    await expect(
      session.callTool(session.credentials, { ...call, id: 'third' }),
    ).resolves.toMatchObject({ output: 'tool_session_call_limit', isError: true })
    expect(() => session.listTools(session.credentials)).toThrow('tool_session_closed')
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
        authorization: {
          host: 'docs',
          policy,
          declaration: { capabilities: ['semantic-read', 'transaction-proposal'] },
        },
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
