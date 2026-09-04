import { describe, expect, it, vi } from 'vitest'
import {
  createPcHostRegistration,
  ENHANCED_HOSTS,
  PC_HOST_CODEX_CHANNELS,
} from '@wiswork/agent-runtime'
import { createSlidesSkill } from '../../slides/src/renderer/ai/slides-skill'
import { ShellCodexRuntime } from '../src/main/codex-runtime'
import { registerPcCodexHosts } from '../src/main/pc-codex-hosts'

describe('PC Codex host registrar', () => {
  it('registers the complete production Slides tool catalog', async () => {
    const handlers = new Map<string, (...args: any[]) => any>()
    const owner = { id: 71, isDestroyed: () => false, send: vi.fn() }
    const registerDocument = vi.fn((input) => {
      expect(input.session.listTools(input.session.credentials).length).toBeGreaterThan(0)
      return () => undefined
    })
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
      bootstrap: {
        start: async () => ({
          registerDocument,
          startTurn: vi.fn(async () => undefined),
          cancelTurn: vi.fn(async () => undefined),
          closeDocument: vi.fn(async () => undefined),
          close: vi.fn(async () => undefined),
        }),
      },
    })
    await runtime.initialize()
    const registrar = registerPcCodexHosts({
      ipcMain: { handle: (channel, handler) => handlers.set(channel, handler) },
      runtime,
      policy,
      hostForOwner: (candidate) => (candidate === owner ? 'slides' : null),
    })
    const skill = createSlidesSkill({
      getSlides: () => [],
      getCurrent: () => 0,
      getSelectedIds: () => [],
      applySlide: () => undefined,
      applyDeck: () => undefined,
      fitWidthPx: 1280,
    })
    const registration = createPcHostRegistration({
      host: 'slides',
      documentId: 'production-slides-catalog',
      generation: 1,
      skill,
    })

    await expect(
      handlers.get(PC_HOST_CODEX_CHANNELS.register)!({ sender: owner }, registration),
    ).resolves.toBeUndefined()
    expect(registerDocument).toHaveBeenCalledOnce()
    await registrar.close()
  })

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
    const telemetry = { host: vi.fn(), component: vi.fn() }
    const registrar = registerPcCodexHosts({
      ipcMain: { handle: (channel, handler) => handlers.set(channel, handler) },
      runtime,
      policy,
      hostForOwner: (candidate) => (candidate === owner ? 'docs' : null),
      telemetry,
    })
    await handlers.get(PC_HOST_CODEX_CHANNELS.telemetry)!(
      { sender: owner },
      { kind: 'host', host: 'docs', phase: 'plan', outcome: 'succeeded' },
    )
    expect(telemetry.host).toHaveBeenCalledWith('docs', 'plan', 'succeeded')
    expect(() =>
      handlers.get(PC_HOST_CODEX_CHANNELS.telemetry)!(
        { sender: owner },
        { kind: 'host', host: 'slides', phase: 'plan', outcome: 'succeeded' },
      ),
    ).toThrow('enhanced_untrusted_request')
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
    expect(
      registered.summarizeProposal({ id: 'x', name: 'replace_blocks', input: { text: 'secret' } }),
    ).toEqual({ operation: 'replace', target: 'blocks', scope: 'bounded-set' })
    expect(registered.summarizeProposal({ id: 'x', name: 'unknown', input: {} })).toBeUndefined()
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

    const readWithoutSnapshotAuthority = registered.session.callTool(
      registered.session.credentials,
      {
        id: 'r2',
        name: 'read_blocks',
        input: {},
      },
    ) as Promise<unknown>
    expect(() =>
      handlers.get(PC_HOST_CODEX_CHANNELS.toolResult)!(
        { sender: owner },
        {
          documentId: 'doc-1',
          generation: 3,
          callId: 'r2',
          execution: { output: 'ok', summary: 'read' },
          snapshotBefore: '123e4567-e89b-42d3-a456-426614174000',
        },
      ),
    ).toThrow('enhanced_untrusted_request')
    await handlers.get(PC_HOST_CODEX_CHANNELS.toolResult)!(
      { sender: owner },
      {
        documentId: 'doc-1',
        generation: 3,
        callId: 'r2',
        execution: { output: 'ok', summary: 'read' },
      },
    )
    await expect(readWithoutSnapshotAuthority).resolves.toMatchObject({ output: 'ok' })

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
      summary: { operation: 'replace', target: 'blocks', scope: 'bounded-set', count: 1 },
    })
    expect(sent.at(-1)).toEqual([
      PC_HOST_CODEX_CHANNELS.proposal,
      expect.objectContaining({
        proposalId: 'P'.repeat(43),
        documentId: 'doc-1',
        generation: 3,
        toolName: 'replace_blocks',
        summary: { operation: 'replace', target: 'blocks', scope: 'bounded-set', count: 1 },
      }),
    ])
    expect(JSON.stringify(sent.at(-1)?.[1])).not.toContain('not-for-ui')
    expect(sent.filter(([channel]) => channel === PC_HOST_CODEX_CHANNELS.toolCall)).toHaveLength(2)
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
        snapshotBefore: '123e4567-e89b-42d3-a456-426614174000',
      },
    )
    await confirmation
    await expect(suspended.result).resolves.toMatchObject({ output: 'applied', mutated: true })
    expect(sent).toContainEqual([
      PC_HOST_CODEX_CHANNELS.event,
      expect.objectContaining({
        type: 'tool-executed',
        event: expect.objectContaining({
          snapshotBefore: '123e4567-e89b-42d3-a456-426614174000',
        }),
      }),
    ])
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
      summary: { operation: 'replace', target: 'blocks', scope: 'bounded-set', count: 1 },
    })
    await handlers.get(PC_HOST_CODEX_CHANNELS.cancelProposal)!(
      { sender: owner },
      'doc-1',
      3,
      'Q'.repeat(43),
    )
    await expect(cancelled.result).resolves.toMatchObject({ isError: true, mutated: false })
    const expired = registered.session.callTool(registered.session.credentials, {
      id: 'm3',
      name: 'replace_blocks',
      input: {},
    }) as any
    registered.onEvent({
      type: 'proposal',
      proposalId: 'R'.repeat(43),
      call: { id: 'm3', name: 'replace_blocks', input: {} },
      expiresAt: Date.now() - 1,
      summary: { operation: 'replace', target: 'blocks', scope: 'bounded-set' },
    })
    expect(() =>
      handlers.get(PC_HOST_CODEX_CHANNELS.confirmProposal)!(
        { sender: owner },
        'doc-1',
        3,
        'R'.repeat(43),
      ),
    ).toThrow('enhanced_proposal_expired')
    await expect(expired.result).resolves.toMatchObject({
      output: 'mutation_expired',
      isError: true,
      mutated: false,
    })
    expect(sent.filter(([channel]) => channel === PC_HOST_CODEX_CHANNELS.toolCall)).toHaveLength(3)
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
