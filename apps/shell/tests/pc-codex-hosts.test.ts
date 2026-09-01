import { describe, expect, it, vi } from 'vitest'
import { ENHANCED_HOSTS, PC_HOST_CODEX_CHANNELS } from '@wiswork/agent-runtime'
import { ShellCodexRuntime } from '../src/main/codex-runtime'
import { registerPcCodexHosts } from '../src/main/pc-codex-hosts'

describe('PC Codex host registrar', () => {
  it('binds registration to the trusted host and dispatches reads and mutations to its owner', async () => {
    const handlers = new Map<string, (...args: any[]) => any>()
    const sent: Array<[string, unknown]> = []
    const owner = {
      id: 7,
      isDestroyed: () => false,
      send: (c: string, v: unknown) => sent.push([c, v]),
    }
    let registered: any
    const engine = {
      registerDocument: vi.fn((input) => {
        registered = input
        return () => undefined
      }),
      startTurn: vi.fn(async () => undefined),
      cancelTurn: vi.fn(async () => undefined),
      closeDocument: vi.fn(async () => undefined),
      close: vi.fn(async () => undefined),
    }
    const policy = {
      globalEnabled: true,
      rawOfficeEnabled: false,
      hosts: Object.fromEntries(ENHANCED_HOSTS.map((host) => [host, true])) as any,
    }
    const runtime = new ShellCodexRuntime({
      activeAgentRuntime: 'enhanced',
      policy,
      isSignedIn: async () => true,
      resolveExecutable: async () => '/private/codex',
      bootstrap: { start: async () => engine },
    })
    await runtime.initialize()
    const registrar = registerPcCodexHosts({
      ipcMain: { handle: (channel, handler) => handlers.set(channel, handler) },
      runtime,
      policy,
      hostForOwner: (candidate) => (candidate === owner ? 'docs' : null),
    })
    await handlers.get(PC_HOST_CODEX_CHANNELS.register)!(
      { sender: owner },
      {
        host: 'docs',
        documentId: 'doc-1',
        generation: 3,
        systemPrompt: 'docs rules',
        tools: [
          { name: 'read_blocks', description: 'read', inputSchema: { type: 'object' } },
          { name: 'replace_blocks', description: 'replace', inputSchema: { type: 'object' } },
        ],
        mutatingTools: ['replace_blocks'],
      },
    )
    const read = registered.session.callTool(registered.session.credentials, {
      id: 'r1',
      name: 'read_blocks',
      input: {},
    }) as Promise<unknown>
    expect(sent.at(-1)?.[0]).toBe(PC_HOST_CODEX_CHANNELS.toolCall)
    await handlers.get(PC_HOST_CODEX_CHANNELS.toolResult)!(
      { sender: owner },
      {
        documentId: 'doc-1',
        generation: 3,
        callId: 'r1',
        execution: { output: 'ok', summary: 'read' },
      },
    )
    await expect(read).resolves.toMatchObject({ output: 'ok' })

    const suspended = registered.session.callTool(registered.session.credentials, {
      id: 'm1',
      name: 'replace_blocks',
      input: {},
    }) as any
    expect(suspended.kind).toBe('tool-execution-suspension')
    registered.onEvent({
      type: 'proposal',
      proposalId: 'P'.repeat(43),
      call: { id: 'm1', name: 'replace_blocks', input: { secret: 'not-for-ui' } },
      expiresAt: Date.now() + 60_000,
    })
    expect(sent.at(-1)).toEqual([
      PC_HOST_CODEX_CHANNELS.proposal,
      expect.objectContaining({
        proposalId: 'P'.repeat(43),
        documentId: 'doc-1',
        generation: 3,
        toolName: 'replace_blocks',
      }),
    ])
    expect(JSON.stringify(sent.at(-1)?.[1])).not.toContain('not-for-ui')
    expect(sent.filter(([channel]) => channel === PC_HOST_CODEX_CHANNELS.toolCall)).toHaveLength(1)
    expect(() =>
      handlers.get(PC_HOST_CODEX_CHANNELS.confirmProposal)!(
        { sender: owner },
        'cross-document',
        3,
        'P'.repeat(43),
      ),
    ).toThrow('enhanced_untrusted_request')
    const confirmation = handlers.get(PC_HOST_CODEX_CHANNELS.confirmProposal)!(
      { sender: owner },
      'doc-1',
      3,
      'P'.repeat(43),
    )
    await vi.waitFor(() => expect(sent.at(-1)?.[1]).toMatchObject({ call: { id: 'm1' } }))
    await handlers.get(PC_HOST_CODEX_CHANNELS.toolResult)!(
      { sender: owner },
      {
        documentId: 'doc-1',
        generation: 3,
        callId: 'm1',
        execution: { output: 'applied', summary: 'replace', mutated: true },
      },
    )
    await confirmation
    await expect(suspended.result).resolves.toMatchObject({ output: 'applied', mutated: true })
    expect(() =>
      handlers.get(PC_HOST_CODEX_CHANNELS.confirmProposal)!(
        { sender: owner },
        'doc-1',
        3,
        'P'.repeat(43),
      ),
    ).toThrow('enhanced_untrusted_request')
    const cancelled = registered.session.callTool(registered.session.credentials, {
      id: 'm2',
      name: 'replace_blocks',
      input: {},
    }) as any
    registered.onEvent({
      type: 'proposal',
      proposalId: 'Q'.repeat(43),
      call: { id: 'm2', name: 'replace_blocks', input: {} },
      expiresAt: Date.now() + 60_000,
    })
    await handlers.get(PC_HOST_CODEX_CHANNELS.cancelProposal)!(
      { sender: owner },
      'doc-1',
      3,
      'Q'.repeat(43),
    )
    await expect(cancelled.result).resolves.toMatchObject({ isError: true, mutated: false })
    expect(sent.filter(([channel]) => channel === PC_HOST_CODEX_CHANNELS.toolCall)).toHaveLength(2)
    await registrar.closeOwner(owner)
    expect(engine.closeDocument).toHaveBeenCalledWith('doc-1')
  })

  it('rejects host substitution before registering a document', async () => {
    const handlers = new Map<string, (...args: any[]) => any>()
    const owner = { id: 8, isDestroyed: () => false, send: vi.fn() }
    registerPcCodexHosts({
      ipcMain: { handle: (c, h) => handlers.set(c, h) },
      runtime: { registerDocument: vi.fn() } as any,
      policy: { globalEnabled: true, rawOfficeEnabled: false, hosts: { docs: true } } as any,
      hostForOwner: () => 'docs',
    })
    await expect(
      handlers.get(PC_HOST_CODEX_CHANNELS.register)!(
        { sender: owner },
        {
          host: 'slides',
          documentId: 'doc',
          generation: 0,
          systemPrompt: '',
          tools: [],
          mutatingTools: [],
        },
      ),
    ).rejects.toThrow('enhanced_untrusted_request')
  })
})
