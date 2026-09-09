import {
  createToolExecutionSuspensionAuthority,
  isToolExecutionSuspension,
  type ToolExecution,
} from '@wiswork/agent-core'
import { describe, expect, it, vi } from 'vitest'
import { createDocumentCarrierIssuer } from '../src/index.js'
import {
  compiledDocumentTool,
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
function policyGrant(host: 'docs' | 'slides' = 'docs') {
  const grant = Object.freeze({})
  const snapshot = {
    generation: 4,
    host,
    policy: rollout,
    capabilities: ['semantic-read', 'transaction-proposal'],
  } as const
  return {
    policyGrant: grant,
    consumePolicyGrant: (candidate: unknown) => {
      if (candidate !== grant) throw new Error('invalid_enhanced_policy_handle')
      return snapshot
    },
  }
}

function fixture(overrides: Partial<DocumentToolRegistration> = {}) {
  const suspensionAuthority = createToolExecutionSuspensionAuthority()
  const executeRead = vi.fn(async (): Promise<ToolExecution> => ({
    output: 'text',
    summary: 'read',
  }))
  let open = true
  const manifest = createDocumentToolManifest({
    ...policyGrant(),
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
    suspendMutation: suspensionAuthority.suspend,
    ownsSuspension: suspensionAuthority.owns,
    ...overrides,
  }
  const session = createDocumentToolSession(registration)
  return { session, executeRead, closeHost: () => (open = false) }
}

describe('document-scoped tool session', () => {
  it('does not expose mutable references to the compiled policy catalog', () => {
    const compiled = compiledDocumentTool('office-powerpoint', 'execute_office_js')!
    expect(() => {
      ;(compiled as unknown as string[])[0] = 'read'
    }).toThrow()
    expect(compiledDocumentTool('office-powerpoint', 'execute_office_js')).toEqual([
      'mutate',
      'transaction-proposal',
    ])
  })

  it('permits PC-backed retrieval tools for the Office PowerPoint host', () => {
    const grant = Object.freeze({})
    expect(() =>
      createDocumentToolManifest({
        policyGrant: grant,
        consumePolicyGrant(candidate: unknown) {
          if (candidate !== grant) throw new Error('invalid_enhanced_policy_handle')
          return {
            generation: 4,
            host: 'office-powerpoint',
            policy: rollout,
            capabilities: ['semantic-read'],
          }
        },
        tools: ['web_search', 'image_search', 'get_presentation_state', 'ask_clarification'].map(
          (name) => ({
            name,
            description: name,
            inputSchema: { type: 'object' },
          }),
        ),
        policy: {
          web_search: 'read',
          image_search: 'read',
          get_presentation_state: 'read',
          ask_clarification: 'read',
        },
      }),
    ).not.toThrow()
  })

  it('compiles the complete Office PowerPoint attachment, media, and editing tool set', () => {
    const expected = {
      read: 'read',
      bash: 'read',
      'insert-image': 'mutate',
      insert_web_image: 'mutate',
      set_slide_background: 'mutate',
    } as const

    for (const [name, mutability] of Object.entries(expected))
      expect(compiledDocumentTool('office-powerpoint', name)?.[0]).toBe(mutability)
  })

  it('separates bounded Office operations from elevated raw Office proposals', () => {
    const bounded = [
      ['office-word', 'execute_office_js'],
      ['office-excel', 'eval_officejs'],
      ['office-powerpoint', 'execute_office_js'],
    ] as const

    for (const [host, name] of bounded) {
      const grant = Object.freeze({})
      expect(() =>
        createDocumentToolManifest({
          policyGrant: grant,
          consumePolicyGrant(candidate: unknown) {
            if (candidate !== grant) throw new Error('invalid_enhanced_policy_handle')
            return {
              generation: 4,
              host,
              policy: rollout,
              capabilities: ['transaction-proposal'] as const,
            }
          },
          tools: [{ name, description: 'Bounded declarative operations.', inputSchema: {} }],
          policy: { [name]: 'mutate' },
        }),
      ).not.toThrow()
    }

    const rawGrant = Object.freeze({})
    expect(() =>
      createDocumentToolManifest({
        policyGrant: rawGrant,
        consumePolicyGrant(candidate: unknown) {
          if (candidate !== rawGrant) throw new Error('invalid_enhanced_policy_handle')
          return {
            generation: 4,
            host: 'office-powerpoint',
            policy: rollout,
            capabilities: ['transaction-proposal'] as const,
          }
        },
        tools: [
          {
            name: 'propose_raw_office_edit',
            description: 'Elevated raw Office proposal.',
            inputSchema: {},
          },
        ],
        policy: { propose_raw_office_edit: 'mutate' },
      }),
    ).toThrow('tool_capability_denied')
  })

  it('waits for human questionnaire answers beyond the ordinary read timeout', async () => {
    vi.useFakeTimers()
    let answer!: (value: ToolExecution) => void
    const f = fixture({
      identity: {
        ownerId: 'owner',
        host: 'slides',
        documentId: 'doc',
        sessionId: 'session_1',
        generation: 4,
      },
      manifest: createDocumentToolManifest({
        ...policyGrant('slides'),
        tools: [{ ...readTool, name: 'ask_clarification' }],
        policy: { ask_clarification: 'read' },
      }),
      executeRead: () =>
        new Promise((resolve) => {
          answer = resolve
        }),
    })
    try {
      let settled = false
      const pending = Promise.resolve(
        f.session.callTool(f.session.credentials, {
          id: 'survey',
          name: 'ask_clarification',
          input: {},
        }),
      ).then((result) => {
        settled = true
        return result
      })
      await vi.advanceTimersByTimeAsync(90_000)
      expect(settled).toBe(false)
      answer({ output: 'answers', summary: 'answered', mutated: false })
      await expect(pending).resolves.toMatchObject({ output: 'answers' })
    } finally {
      f.session.close()
      vi.useRealTimers()
    }
  })
  it('does not claim another proposal when the requested consent already expired', async () => {
    vi.useFakeTimers()
    const f = fixture()
    try {
      f.session.callTool(f.session.credentials, { id: 'first', name: writeTool.name, input: {} })
      await vi.advanceTimersByTimeAsync(10_000)
      f.session.callTool(f.session.credentials, { id: 'second', name: writeTool.name, input: {} })
      await vi.advanceTimersByTimeAsync(290_000)
      expect(f.session.mutationAuthority.claimNext('first')).toBeUndefined()
      expect(f.session.mutationAuthority.claimNext('second')?.request.call.id).toBe('second')
    } finally {
      f.session.close()
      vi.useRealTimers()
    }
  })
  it('keeps consent queued beyond 30 seconds and starts execution timeout only when claimed', async () => {
    vi.useFakeTimers()
    const f = fixture()
    try {
      const outcome = f.session.callTool(f.session.credentials, {
        id: 'wait',
        name: writeTool.name,
        input: {},
      }) as any
      const finished = vi.fn()
      void outcome.result.then(finished)
      await vi.advanceTimersByTimeAsync(240_000)
      expect(finished).not.toHaveBeenCalled()
      const claimed = f.session.mutationAuthority.claimNext()!
      expect(claimed).toBeDefined()
      await vi.advanceTimersByTimeAsync(29_999)
      expect(finished).not.toHaveBeenCalled()
      f.session.mutationAuthority.settle(claimed.claim, {
        output: 'applied',
        summary: 'applied',
        mutated: true,
      })
      await expect(outcome.result).resolves.toMatchObject({ output: 'applied', mutated: true })
      expect(() =>
        f.session.mutationAuthority.settle(claimed.claim, { output: 'again', summary: 'again' }),
      ).toThrow('mutation_claim_consumed')
    } finally {
      f.session.close()
      vi.useRealTimers()
    }
  })

  it('expires unclaimed consent at five minutes without allowing a later claim', async () => {
    vi.useFakeTimers()
    const f = fixture()
    try {
      const outcome = f.session.callTool(f.session.credentials, {
        id: 'expire',
        name: writeTool.name,
        input: {},
      }) as any
      await vi.advanceTimersByTimeAsync(300_000)
      await expect(outcome.result).resolves.toMatchObject({
        output: 'mutation_expired',
        isError: true,
        mutated: false,
      })
      expect(f.session.mutationAuthority.claimNext()).toBeUndefined()
    } finally {
      f.session.close()
      vi.useRealTimers()
    }
  })

  it('still times out claimed execution after 30 seconds and rejects late receipts', async () => {
    vi.useFakeTimers()
    const f = fixture()
    try {
      const outcome = f.session.callTool(f.session.credentials, {
        id: 'slow-write',
        name: writeTool.name,
        input: {},
      }) as any
      await vi.advanceTimersByTimeAsync(40_000)
      const claimed = f.session.mutationAuthority.claimNext()!
      expect(claimed).toBeDefined()
      await vi.advanceTimersByTimeAsync(30_000)
      await expect(outcome.result).resolves.toMatchObject({ output: 'tool_timeout', isError: true })
      expect(() =>
        f.session.mutationAuthority.settle(claimed.claim, { output: 'late', summary: 'late' }),
      ).toThrow('mutation_claim_consumed')
    } finally {
      f.session.close()
      vi.useRealTimers()
    }
  })

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

  it('routes reads to executeRead and queues detached mutations without a writer callback', async () => {
    const f = fixture()
    await expect(
      f.session.callTool(f.session.credentials, { id: 'r', name: readTool.name, input: {} }),
    ).resolves.toMatchObject({ output: 'text' })
    const outcome = f.session.callTool(f.session.credentials, {
      id: 'w',
      name: writeTool.name,
      input: {},
    })
    expect(isToolExecutionSuspension(outcome as any)).toBe(true)
    const claimed = f.session.mutationAuthority.claimNext()!
    expect(claimed.request).toMatchObject({
      identity: { documentId: 'doc', generation: 4 },
      call: { id: 'w', name: writeTool.name },
      catalogDigest: f.session.catalogDigest,
    })
    f.session.mutationAuthority.settle(claimed.claim, {
      output: 'changed',
      summary: 'changed',
      mutated: true,
    })
    await expect((outcome as any).result).resolves.toMatchObject({
      output: 'changed',
      mutated: true,
    })
    expect(f.executeRead).toHaveBeenCalledTimes(1)
  })

  it('rejects a mutation suspension minted by a different owner', async () => {
    const owner = createToolExecutionSuspensionAuthority()
    const attacker = createToolExecutionSuspensionAuthority()
    const f = fixture({
      suspendMutation: attacker.suspend,
      ownsSuspension: owner.owns,
    })
    const outcome = await f.session.callTool(f.session.credentials, {
      id: 'foreign',
      name: 'replace_blocks',
      input: {},
    })
    expect(isToolExecutionSuspension(outcome as any)).toBe(false)
    expect(outcome).toMatchObject({ output: 'tool_authority_denied', isError: true })
    expect(f.session.mutationAuthority.claimNext()).toBeUndefined()
  })

  it.each([
    [
      'suspension mint failure',
      () => {
        throw new Error('mint failed')
      },
      () => true,
    ],
    [
      'suspension ownership failure',
      (result: Promise<ToolExecution>) => result as any,
      () => {
        throw new Error('ownership failed')
      },
    ],
  ])(
    'releases the pending mutation gate after %s',
    async (_label, suspendMutation, ownsSuspension) => {
      const f = fixture({ suspendMutation: suspendMutation as any, ownsSuspension })
      await expect(
        Promise.resolve(
          f.session.callTool(f.session.credentials, {
            id: 'broken-suspension',
            name: writeTool.name,
            input: {},
          }),
        ),
      ).resolves.toMatchObject({ output: 'tool_authority_denied', isError: true })
      await expect(
        f.session.callTool(f.session.credentials, {
          id: 'read-after-broken-suspension',
          name: readTool.name,
          input: {},
        }),
      ).resolves.toMatchObject({ output: 'text' })
      expect(f.session.mutationAuthority.claimNext()).toBeUndefined()
    },
  )

  it.each(['success', 'failure', 'timeout', 'cancel'] as const)(
    'releases the pending mutation gate after terminal %s',
    async (terminal) => {
      if (terminal === 'timeout') vi.useFakeTimers()
      const f = fixture({ maxCallMs: 10 })
      try {
        const outcome = f.session.callTool(f.session.credentials, {
          id: `write-${terminal}`,
          name: writeTool.name,
          input: {},
        }) as any
        const claimed = f.session.mutationAuthority.claimNext()!
        if (terminal === 'success')
          f.session.mutationAuthority.settle(claimed.claim, {
            output: 'applied',
            summary: 'applied',
            mutated: true,
          })
        else if (terminal === 'failure')
          f.session.mutationAuthority.reject(claimed.claim, 'apply_failed')
        else if (terminal === 'cancel') f.session.cancel(f.session.credentials, `write-${terminal}`)
        else await vi.advanceTimersByTimeAsync(10)
        await outcome.result
        expect(claimed.request.signal.aborted).toBe(terminal === 'timeout' || terminal === 'cancel')
        await expect(
          f.session.callTool(f.session.credentials, {
            id: `read-after-${terminal}`,
            name: readTool.name,
            input: {},
          }),
        ).resolves.toMatchObject({ output: 'text' })
      } finally {
        f.session.close()
        if (terminal === 'timeout') vi.useRealTimers()
      }
    },
  )

  it('cancels queued mutations by call id before an authority can claim them', async () => {
    const f = fixture()
    const outcome = f.session.callTool(f.session.credentials, {
      id: 'queued-write',
      name: writeTool.name,
      input: {},
    })
    expect(isToolExecutionSuspension(outcome as any)).toBe(true)
    expect(f.session.cancel(f.session.credentials, 'queued-write')).toBe(true)
    await expect((outcome as any).result).resolves.toMatchObject({
      output: 'tool_cancelled',
      isError: true,
    })
    expect(f.session.mutationAuthority.claimNext()).toBeUndefined()
  })

  it('cancels all queued and claimed mutations for a revoked turn', async () => {
    const f = fixture()
    const queued = f.session.callTool(f.session.credentials, {
      id: 'queued-all',
      name: writeTool.name,
      input: {},
    }) as any
    const claimed = f.session.mutationAuthority.claimNext()!
    const second = f.session.callTool(f.session.credentials, {
      id: 'second-all',
      name: writeTool.name,
      input: {},
    }) as any
    expect(f.session.cancelAll(f.session.credentials)).toBe(2)
    await expect(queued.result).resolves.toMatchObject({ output: 'tool_cancelled', isError: true })
    await expect(second.result).resolves.toMatchObject({ output: 'tool_cancelled', isError: true })
    expect(() =>
      f.session.mutationAuthority.settle(claimed.claim, { output: 'late', summary: 'late' }),
    ).toThrow('mutation_claim_consumed')
  })

  it('rejects cross-session, unknown, oversized, cancelled and closed calls', async () => {
    const a = fixture(),
      b = fixture()
    expect(() =>
      a.session.callTool(b.session.credentials, { id: 'x', name: readTool.name, input: {} }),
    ).toThrow('tool_unauthorized')
    expect(
      a.session.callTool(a.session.credentials, { id: 'unknown', name: 'shell', input: {} }),
    ).toMatchObject({ output: 'unknown_tool', isError: true })
    expect(
      a.session.callTool(a.session.credentials, {
        id: 'big',
        name: readTool.name,
        input: { text: 'x'.repeat(1_000_001) },
      }),
    ).toMatchObject({ output: 'invalid_tool_call', isError: true })
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

  it('bounds read execution and rejects read mutation policy violations', async () => {
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

  it('allows a model tool batch to run multiple bounded reads concurrently', async () => {
    let releaseFirst!: () => void
    const firstPending = new Promise<void>((resolve) => {
      releaseFirst = resolve
    })
    const executeRead = vi.fn(async (call: { id: string }): Promise<ToolExecution> => {
      if (call.id === 'first') await firstPending
      return { output: call.id, summary: call.id, mutated: false }
    })
    const f = fixture({ executeRead })

    const first = f.session.callTool(f.session.credentials, {
      id: 'first',
      name: readTool.name,
      input: {},
    }) as Promise<ToolExecution>
    const second = f.session.callTool(f.session.credentials, {
      id: 'second',
      name: readTool.name,
      input: {},
    }) as Promise<ToolExecution>

    await expect(Promise.resolve(second)).resolves.toMatchObject({ output: 'second' })
    expect(executeRead).toHaveBeenCalledTimes(2)
    releaseFirst()
    await expect(first).resolves.toMatchObject({ output: 'first' })
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
