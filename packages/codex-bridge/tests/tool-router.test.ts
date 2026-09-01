import { suspendToolExecution, type ToolExecution } from '@wiswork/agent-core'
import { describe, expect, it, vi } from 'vitest'
import { createDocumentCarrierIssuer } from '../src/index.js'
import {
  createDocumentToolManifest,
  createDocumentToolSession,
  ToolRouterError,
  type DocumentToolRegistration,
} from '../src/tool-router.js'
import captured from './fixtures/codex-0147-request.json'

const rollout = {
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
const readTool = {
  name: 'get_document_context',
  description: 'Read.',
  inputSchema: { type: 'object' },
}
const writeTool = { name: 'replace_blocks', description: 'Write.', inputSchema: { type: 'object' } }

function fixture(overrides: Partial<DocumentToolRegistration> = {}) {
  const executeRead = vi.fn(async (): Promise<ToolExecution> => ({
    output: 'text',
    summary: 'read',
  }))
  const prepareMutation = vi.fn(() =>
    suspendToolExecution(Promise.resolve({ output: 'changed', summary: 'changed', mutated: true })),
  )
  let open = true
  const manifest = createDocumentToolManifest({
    authorization: {
      host: 'docs',
      policy: rollout,
      declaration: { capabilities: ['semantic-read', 'transaction-proposal'] },
    },
    tools: [readTool, writeTool],
    policy: { get_document_context: 'read', replace_blocks: 'mutate' },
  })
  const registration: DocumentToolRegistration = {
    identity: {
      ownerId: 'owner',
      host: 'docs',
      documentId: 'doc',
      sessionId: 'session_1',
      generation: 4,
    },
    manifest,
    isOpen: () => open,
    executeRead,
    prepareMutation,
    ...overrides,
  }
  const session = createDocumentToolSession(registration)
  return { session, executeRead, prepareMutation, closeHost: () => (open = false) }
}

describe('document-scoped tool session', () => {
  it('binds canonical high-entropy credentials and immutable exact identity', () => {
    const f = fixture()
    expect(Buffer.from(f.session.credentials.sessionId, 'base64url')).toHaveLength(32)
    expect(Buffer.from(f.session.credentials.secret, 'base64url')).toHaveLength(32)
    expect(Object.isFrozen(f.session.identity)).toBe(true)
    expect(() =>
      f.session.authorize({ ...f.session.credentials, secret: `${f.session.credentials.secret}=` }),
    ).toThrow('tool_unauthorized')
    expect(() =>
      f.session.authorize({ ...f.session.credentials, sessionId: 'A'.repeat(43) }),
    ).toThrow('tool_unauthorized')
  })

  it('routes reads only to executeRead and mutations only to prepareMutation', async () => {
    const f = fixture()
    await expect(
      f.session.callTool(f.session.credentials, { id: 'r', name: readTool.name, input: {} }),
    ).resolves.toMatchObject({ output: 'text' })
    expect(f.prepareMutation).not.toHaveBeenCalled()
    await expect(
      f.session.callTool(f.session.credentials, { id: 'w', name: writeTool.name, input: {} }),
    ).resolves.toMatchObject({ output: 'changed', mutated: true })
    expect(f.executeRead).toHaveBeenCalledTimes(1)
  })

  it('rejects cross-session, unknown, oversized, cancelled and closed calls', async () => {
    const a = fixture(),
      b = fixture()
    await expect(
      a.session.callTool(b.session.credentials, { id: 'x', name: readTool.name, input: {} }),
    ).rejects.toThrow('tool_unauthorized')
    await expect(
      a.session.callTool(a.session.credentials, { id: 'unknown', name: 'shell', input: {} }),
    ).resolves.toMatchObject({ output: 'unknown_tool', isError: true })
    await expect(
      a.session.callTool(a.session.credentials, {
        id: 'big',
        name: readTool.name,
        input: { text: 'x'.repeat(1_000_001) },
      }),
    ).resolves.toMatchObject({ output: 'invalid_tool_call', isError: true })
    const controller = new AbortController()
    controller.abort()
    await expect(
      a.session.callTool(
        a.session.credentials,
        { id: 'cancel', name: readTool.name, input: {} },
        controller.signal,
      ),
    ).resolves.toMatchObject({ output: 'tool_cancelled', isError: true })
    a.session.close()
    expect(() => a.session.listTools(a.session.credentials)).toThrow('tool_session_closed')
  })

  it('bounds execution and rejects mutation result or read mutation policy violations', async () => {
    const mutatingRead = fixture({
      executeRead: vi.fn(async () => ({ output: 'bad', summary: 'bad', mutated: true })),
    })
    await expect(
      mutatingRead.session.callTool(mutatingRead.session.credentials, {
        id: 'r',
        name: readTool.name,
        input: {},
      }),
    ).resolves.toMatchObject({ output: 'tool_policy_violation' })
    const direct = fixture({
      prepareMutation: vi.fn(async () => ({ output: 'bad', summary: 'bad', mutated: true }) as any),
    })
    await expect(
      direct.session.callTool(direct.session.credentials, {
        id: 'w',
        name: writeTool.name,
        input: {},
      }),
    ).resolves.toMatchObject({ output: 'mutation_authority_required' })
    const slow = fixture({
      maxCallMs: 10,
      executeRead: vi.fn(async () => await new Promise<ToolExecution>(() => undefined)),
    })
    await expect(
      slow.session.callTool(slow.session.credentials, {
        id: 'slow',
        name: readTool.name,
        input: {},
      }),
    ).resolves.toMatchObject({ output: 'tool_timeout' })
  })

  it('binds the catalog digest into a Task 2 one-use carrier', () => {
    const validate = vi.fn((capability: unknown) => capability === 'opaque')
    const issuer = createDocumentCarrierIssuer(
      { host: 'docs', documentId: 'doc', sessionId: 'session_1', generation: 4 },
      validate,
    )
    const f = fixture({ carrier: { issuer, capability: 'opaque' } })
    const handle = f.session.issueCarrier(f.session.credentials, {
      turnId: 'turn_1',
      sourceNonce: 'N'.repeat(43),
      toolName: 'get_document_context',
    })
    const body = structuredClone(captured) as any
    const metadata = JSON.parse(body.client_metadata['x-codex-turn-metadata'])
    metadata.code_mode_tool_names = {
      mcp__wiswork__get_document_context: {
        name: 'get_document_context',
        namespace: 'mcp__wiswork',
      },
    }
    body.client_metadata['x-codex-turn-metadata'] = JSON.stringify(metadata)
    const developer = body.input.find((item: any) => item.type === 'additional_tools')
    const exec = developer.tools
      .find((item: any) => item.name === 'functions')
      .tools.find((item: any) => item.name === 'exec')
    exec.description = `Execute exactly one document MCP call. An optional first line // @exec: {"yield_time_ms":1000,"max_output_tokens":100} is allowed. Allowed syntax: text(await tools.mcp__wiswork__get_document_context({...})). Arguments must be a JSON object literal. No other JavaScript is allowed.`
    expect(validate).toHaveBeenCalledWith(
      'opaque',
      expect.objectContaining({
        schemaDigest: f.session.catalogDigest,
        documentId: 'doc',
        generation: 4,
      }),
    )
    expect(() => issuer.prepareTurn(body, {}, handle)).not.toThrow()
    expect(() => issuer.prepareTurn(body, {}, handle)).toThrow('carrier_authorization_consumed')
  })

  it('rejects forged manifests and invalid exact policy', () => {
    const f = fixture()
    expect(() =>
      createDocumentToolSession({
        ...({} as any),
        ...f,
        manifest: { digest: f.session.catalogDigest } as any,
      }),
    ).toThrowError(ToolRouterError)
  })
})
