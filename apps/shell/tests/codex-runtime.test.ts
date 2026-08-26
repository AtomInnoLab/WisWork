import { EventEmitter } from 'node:events'
import { describe, expect, it, vi } from 'vitest'
import type {
  CodexAppServerClient,
  CodexAppServerNotification,
  CodexProcessManager,
  DocumentMcpServer,
  MessagesRequest,
  ResponsesBridge,
  ToolSessionRegistration,
} from '@wiswork/codex-bridge'
import { CODEX_RUNTIME_CHANNELS } from '../src/shared/codex-api'
import {
  ShellCodexRuntime,
  ShellCodexRuntimeError,
  WISWORK_CODEX_DEVELOPER_INSTRUCTIONS,
  createCodexBeforeQuitHandler,
  createWisUsageBridgeFetch,
  logoutWithCodexClose,
  prepareCodexDocumentClose,
  prepareCodexDocumentsClose,
  runCodexPreparedClose,
} from '../src/main/codex-runtime'

class FakeOwner extends EventEmitter {
  readonly sent: Array<{ channel: string; payload: unknown }> = []
  destroyed = false

  isDestroyed(): boolean {
    return this.destroyed
  }

  send(channel: string, payload: unknown): void {
    if (this.destroyed) throw new Error('destroyed')
    this.sent.push({ channel, payload })
  }
}

class FakeClient {
  readonly calls: string[]
  readonly listeners = new Set<(notification: CodexAppServerNotification) => void>()
  startTurnError: Error | undefined
  startTurnResult: Promise<{ turn: { id: string } }> | undefined
  interruptTurnError: Error | undefined

  constructor(calls: string[]) {
    this.calls = calls
  }

  async initialize(): Promise<Record<string, string>> {
    this.calls.push('client.initialize')
    return {
      userAgent: 'codex',
      codexHome: '/tmp/home',
      platformFamily: 'unix',
      platformOs: 'linux',
    }
  }

  async startThread(): Promise<{ thread: { id: string } }> {
    this.calls.push('client.startThread')
    return { thread: { id: 'thread-1' } }
  }

  async startTurn(): Promise<{ turn: { id: string } }> {
    this.calls.push('client.startTurn')
    if (this.startTurnError) throw this.startTurnError
    if (this.startTurnResult) return this.startTurnResult
    return { turn: { id: 'turn-1' } }
  }

  async interruptTurn(threadId: string, turnId: string): Promise<Record<string, never>> {
    this.calls.push(`client.interrupt:${threadId}:${turnId}`)
    if (this.interruptTurnError) throw this.interruptTurnError
    return {}
  }

  onNotification(listener: (notification: CodexAppServerNotification) => void): () => void {
    this.listeners.add(listener)
    return () => this.listeners.delete(listener)
  }

  emit(notification: CodexAppServerNotification): void {
    for (const listener of [...this.listeners]) listener(notification)
  }
}

interface Deferred<T> {
  promise: Promise<T>
  resolve(value: T): void
  reject(error: unknown): void
}

function deferred<T>(): Deferred<T> {
  let resolve!: (value: T) => void
  let reject!: (error: unknown) => void
  const promise = new Promise<T>((done, fail) => {
    resolve = done
    reject = fail
  })
  return { promise, resolve, reject }
}

function registration(): ToolSessionRegistration {
  return {
    skill: {
      id: 'docs',
      systemPrompt: 'Use document tools.',
      tools: [
        {
          name: 'read_document',
          description: 'Read the document.',
          inputSchema: { type: 'object', additionalProperties: false },
        },
      ],
      executeTool: () => ({ output: 'document', summary: 'Read document' }),
    },
    policy: { read_document: 'read' },
    isOpen: () => true,
    getRevision: () => 'rev-1',
    requestApproval: () => false,
    captureSnapshot: () => 'snapshot-1',
  }
}

function harness(options: { loggedIn?: boolean; deferManagerStart?: boolean } = {}) {
  const calls: string[] = []
  const owner = new FakeOwner()
  const client = new FakeClient(calls)
  const crash = deferred<Error>()
  const managerStart = deferred<CodexAppServerClient>()
  const bridge = {
    baseUrl: 'http://127.0.0.1:41001',
    responsesUrl: 'http://127.0.0.1:41001/v1/responses',
    secret: 'bridge-secret',
    close: vi.fn(async () => {
      calls.push('bridge.close')
    }),
  } satisfies ResponsesBridge
  const mcpSession = {
    url: `http://127.0.0.1:41002/mcp/${'s'.repeat(43)}`,
    secret: 't'.repeat(43),
    close: vi.fn(() => {
      calls.push('mcpSession.close')
    }),
  }
  const mcp = {
    baseUrl: 'http://127.0.0.1:41002',
    register: vi.fn(() => {
      calls.push('mcp.register')
      return mcpSession
    }),
    close: vi.fn(async () => {
      calls.push('mcp.close')
    }),
  } satisfies DocumentMcpServer
  const manager = {
    crashed: crash.promise,
    start: vi.fn(async () => {
      calls.push('manager.start')
      if (options.deferManagerStart) return managerStart.promise
      return client as unknown as CodexAppServerClient
    }),
    stop: vi.fn(async () => {
      calls.push('manager.stop')
    }),
  }
  const authClient = {
    getValidAccountStatus: vi.fn(async () => {
      calls.push('auth.status')
      return options.loggedIn === false ? { loggedIn: false as const } : { loggedIn: true as const }
    }),
    fetchWithAuth: vi.fn(),
  }
  const onProcessCrash = vi.fn(() => {
    calls.push('rendererToolIpc.closeDocument')
  })
  const runtime = new ShellCodexRuntime({
    runtimeKind: 'codex',
    executablePath: '/opt/wiswork/codex',
    authClient,
    startResponsesBridge: vi.fn(async () => {
      calls.push('bridge.start')
      return bridge
    }),
    startDocumentMcpServer: vi.fn(async () => {
      calls.push('mcp.start')
      return mcp
    }),
    createProcessManager: vi.fn((processOptions) => {
      calls.push('manager.create')
      expect(processOptions.developerInstructions).toBe(WISWORK_CODEX_DEVELOPER_INSTRUCTIONS)
      expect(processOptions.bridge).toBe(bridge)
      expect(processOptions.mcp).toEqual({ url: mcpSession.url, secret: mcpSession.secret })
      return manager as unknown as CodexProcessManager
    }),
    onProcessCrash,
  })
  return {
    runtime,
    owner,
    client,
    manager,
    managerStart,
    crash,
    bridge,
    mcp,
    mcpSession,
    authClient,
    onProcessCrash,
    calls,
  }
}

async function startDocument(f: ReturnType<typeof harness>): Promise<void> {
  await f.runtime.registerDocument({
    documentId: 'doc-1',
    owner: f.owner,
    registration: registration(),
  })
}

async function flush(): Promise<void> {
  await new Promise((resolve) => setTimeout(resolve, 0))
}

describe('ShellCodexRuntime', () => {
  it('resolves the verified optional component before starting any local server', async () => {
    const f = harness()
    const resolveExecutable = vi.fn(async () => {
      f.calls.push('component.resolve')
      return '/private/verified/bin/codex'
    })
    const runtime = new ShellCodexRuntime({
      runtimeKind: 'codex',
      resolveExecutable,
      authClient: f.authClient,
      startResponsesBridge: async () => {
        f.calls.push('bridge.start')
        return f.bridge
      },
      startDocumentMcpServer: async () => {
        f.calls.push('mcp.start')
        return f.mcp
      },
      createProcessManager: (options) => {
        expect(options.executablePath).toBe('/private/verified/bin/codex')
        return f.manager as unknown as CodexProcessManager
      },
    })

    await runtime.registerDocument({
      documentId: 'doc-verified',
      owner: f.owner,
      registration: registration(),
    })
    expect(f.calls.slice(0, 4)).toEqual([
      'auth.status',
      'component.resolve',
      'bridge.start',
      'mcp.start',
    ])
    await runtime.closeDocument('doc-verified')
  })

  it('fails visibly before local startup when Enhanced mode is selected but not installed', async () => {
    const f = harness()
    const runtime = new ShellCodexRuntime({
      runtimeKind: 'codex',
      resolveExecutable: vi.fn(async () => {
        throw new Error('private component detail')
      }),
      authClient: f.authClient,
      startResponsesBridge: async () => f.bridge,
      startDocumentMcpServer: async () => f.mcp,
      createProcessManager: () => f.manager as unknown as CodexProcessManager,
    })

    await expect(
      runtime.registerDocument({
        documentId: 'doc-missing',
        owner: f.owner,
        registration: registration(),
      }),
    ).rejects.toMatchObject({ code: 'enhanced_mode_install_required' })
    expect(f.calls).toEqual(['auth.status'])
  })

  it('starts an authenticated document chain in strict dependency order', async () => {
    const f = harness()
    await startDocument(f)

    expect(f.calls).toEqual([
      'auth.status',
      'bridge.start',
      'mcp.start',
      'mcp.register',
      'manager.create',
      'manager.start',
      'client.initialize',
      'client.startThread',
    ])
  })

  it('fails closed before binding local services when WisPaper auth is absent', async () => {
    const f = harness({ loggedIn: false })

    await expect(
      f.runtime.registerDocument({
        documentId: 'doc-1',
        owner: f.owner,
        registration: registration(),
      }),
    ).rejects.toMatchObject({ code: 'auth_required' })
    expect(f.calls).toEqual(['auth.status'])
  })

  it('rolls back the bridge when MCP startup fails', async () => {
    const f = harness()
    const startMcp = vi.fn(async () => {
      f.calls.push('mcp.start')
      throw new Error('raw mcp failure with token')
    })
    const runtime = new ShellCodexRuntime({
      runtimeKind: 'codex',
      executablePath: '/opt/wiswork/codex',
      authClient: f.authClient,
      startResponsesBridge: async () => {
        f.calls.push('bridge.start')
        return f.bridge
      },
      startDocumentMcpServer: startMcp,
      createProcessManager: () => f.manager as unknown as CodexProcessManager,
    })

    await expect(
      runtime.registerDocument({
        documentId: 'doc-1',
        owner: f.owner,
        registration: registration(),
      }),
    ).rejects.toMatchObject({ code: 'codex_mcp_start_failed' })
    expect(f.calls).toEqual(['auth.status', 'bridge.start', 'mcp.start', 'bridge.close'])
  })

  it('closes process, MCP credentials/server, and bridge on document close', async () => {
    const f = harness()
    await startDocument(f)
    f.calls.length = 0

    await f.runtime.closeDocument('doc-1')
    await f.runtime.closeDocument('doc-1')

    expect(f.calls).toEqual(['manager.stop', 'mcpSession.close', 'mcp.close', 'bridge.close'])
  })

  it('tears down every document on logout and app shutdown idempotently', async () => {
    const f = harness()
    await startDocument(f)
    f.calls.length = 0

    await Promise.all([f.runtime.logout(), f.runtime.shutdown(), f.runtime.shutdown()])

    expect(f.manager.stop).toHaveBeenCalledOnce()
    expect(f.mcpSession.close).toHaveBeenCalledOnce()
    expect(f.mcp.close).toHaveBeenCalledOnce()
    expect(f.bridge.close).toHaveBeenCalledOnce()
  })

  it('interrupts cancellation and reports the matching completed turn as cancelled', async () => {
    const f = harness()
    await startDocument(f)
    await expect(f.runtime.startTurn(f.owner, 'doc-1', 'Summarize this.')).resolves.toEqual({
      turnId: 'turn-1',
    })

    await f.runtime.cancelTurn(f.owner, 'doc-1')
    f.client.emit({
      method: 'turn/completed',
      params: { threadId: 'thread-1', turn: { id: 'turn-1', status: 'interrupted' } },
    })

    expect(f.calls).toContain('client.interrupt:thread-1:turn-1')
    expect(f.owner.sent.at(-1)).toEqual({
      channel: CODEX_RUNTIME_CHANNELS.event,
      payload: {
        documentId: 'doc-1',
        event: {
          type: 'done',
          result: { text: '', cancelled: true, turnLimit: false },
        },
      },
    })
  })

  it('interrupts exactly once when cancellation races the turn/start response', async () => {
    const f = harness()
    await startDocument(f)
    const accepted = deferred<{ turn: { id: string } }>()
    f.client.startTurnResult = accepted.promise
    const starting = f.runtime.startTurn(f.owner, 'doc-1', 'Slow start')
    await flush()

    await f.runtime.cancelTurn(f.owner, 'doc-1')
    f.client.emit({
      method: 'turn/started',
      params: { threadId: 'thread-1', turn: { id: 'turn-1' } },
    })
    accepted.resolve({ turn: { id: 'turn-1' } })

    await expect(starting).resolves.toEqual({ turnId: 'turn-1' })
    expect(f.calls.filter((call) => call === 'client.interrupt:thread-1:turn-1')).toHaveLength(1)
  })

  it('accepts a terminal notification that arrives before the turn/start response', async () => {
    const f = harness()
    await startDocument(f)
    const accepted = deferred<{ turn: { id: string } }>()
    f.client.startTurnResult = accepted.promise
    const starting = f.runtime.startTurn(f.owner, 'doc-1', 'Fast completion')
    await flush()
    f.client.emit({
      method: 'turn/started',
      params: { threadId: 'thread-1', turn: { id: 'turn-1' } },
    })
    f.client.emit({
      method: 'turn/completed',
      params: { threadId: 'thread-1', turn: { id: 'turn-1', status: 'completed' } },
    })
    accepted.resolve({ turn: { id: 'turn-1' } })

    await expect(starting).resolves.toEqual({ turnId: 'turn-1' })
    expect(f.calls.some((call) => call.startsWith('client.interrupt:'))).toBe(false)
  })

  it('normalizes cumulative text and a successful completion for the current turn only', async () => {
    const f = harness()
    await startDocument(f)
    await f.runtime.startTurn(f.owner, 'doc-1', 'Answer.')

    f.client.emit({
      method: 'item/agentMessage/delta',
      params: { threadId: 'other', turnId: 'turn-1', itemId: 'i', delta: 'secret' },
    })
    f.client.emit({
      method: 'item/agentMessage/delta',
      params: { threadId: 'thread-1', turnId: 'turn-1', itemId: 'i', delta: 'hel' },
    })
    f.client.emit({
      method: 'item/agentMessage/delta',
      params: { threadId: 'thread-1', turnId: 'turn-1', itemId: 'i', delta: 'lo' },
    })
    f.client.emit({
      method: 'turn/completed',
      params: { threadId: 'thread-1', turn: { id: 'turn-1', status: 'completed' } },
    })

    expect(f.owner.sent).toEqual([
      {
        channel: CODEX_RUNTIME_CHANNELS.event,
        payload: { documentId: 'doc-1', event: { type: 'text', text: 'hel' } },
      },
      {
        channel: CODEX_RUNTIME_CHANNELS.event,
        payload: { documentId: 'doc-1', event: { type: 'text', text: 'hello' } },
      },
      {
        channel: CODEX_RUNTIME_CHANNELS.event,
        payload: {
          documentId: 'doc-1',
          event: {
            type: 'done',
            result: { text: 'hello', cancelled: false, turnLimit: false },
          },
        },
      },
    ])
  })

  it('normalizes document tool activity without feeding display side-channel data to MCP', async () => {
    const f = harness()
    await startDocument(f)
    await f.runtime.startTurn(f.owner, 'doc-1', 'Read it.')
    const routed = f.mcp.register.mock.calls[0]![0] as ToolSessionRegistration

    await expect(
      Promise.resolve(routed.skill.executeTool({ id: 'call-1', name: 'read_document', input: {} })),
    ).resolves.toEqual({ output: 'document', summary: 'Read document' })

    expect(f.owner.sent).toEqual([
      {
        channel: CODEX_RUNTIME_CHANNELS.event,
        payload: {
          documentId: 'doc-1',
          event: {
            type: 'tool-start',
            call: { id: 'call-1', name: 'read_document', input: {} },
          },
        },
      },
      {
        channel: CODEX_RUNTIME_CHANNELS.event,
        payload: {
          documentId: 'doc-1',
          event: {
            type: 'tool-executed',
            call: { id: 'call-1', name: 'read_document', input: {} },
            execution: { output: 'document', summary: 'Read document' },
          },
        },
      },
      {
        channel: CODEX_RUNTIME_CHANNELS.event,
        payload: { documentId: 'doc-1', event: { type: 'turn-end' } },
      },
    ])
  })

  it('allows exactly one direct document tool call across an app-server turn', async () => {
    const f = harness()
    await startDocument(f)
    await f.runtime.startTurn(f.owner, 'doc-1', 'Read it once.')
    const routed = f.mcp.register.mock.calls[0]![0] as ToolSessionRegistration

    await expect(
      Promise.resolve(routed.skill.executeTool({ id: 'call-1', name: 'read_document', input: {} })),
    ).resolves.toMatchObject({ output: 'document' })
    await expect(async () =>
      routed.skill.executeTool({ id: 'call-2', name: 'read_document', input: {} }),
    ).rejects.toThrow('codex_tool_call_limit')
    await expect(async () =>
      routed.skill.executeTool({ id: 'call-1', name: 'read_document', input: {} }),
    ).rejects.toThrow('codex_tool_call_limit')
  })

  it('surfaces child crashes with a stable error and tears the chain down', async () => {
    const f = harness()
    await startDocument(f)
    f.calls.length = 0

    f.crash.resolve(new Error('stderr contains prompt and bearer token'))
    await flush()

    expect(f.owner.sent.at(-1)).toEqual({
      channel: CODEX_RUNTIME_CHANNELS.event,
      payload: {
        documentId: 'doc-1',
        event: {
          type: 'error',
          code: 'enhanced_mode_stopped',
          message: 'Enhanced mode stopped.',
        },
      },
    })
    expect(f.onProcessCrash).toHaveBeenCalledWith('doc-1')
    expect(f.calls).toEqual([
      'rendererToolIpc.closeDocument',
      'manager.stop',
      'mcpSession.close',
      'mcp.close',
      'bridge.close',
    ])
  })

  it('fails a turn visibly without invoking or silently falling back to legacy', async () => {
    const f = harness()
    await startDocument(f)
    f.client.startTurnError = new Error('raw prompt and secret')

    await expect(f.runtime.startTurn(f.owner, 'doc-1', 'Do not replay me')).rejects.toMatchObject({
      code: 'codex_turn_start_failed',
    })
    expect(f.owner.sent.at(-1)).toEqual({
      channel: CODEX_RUNTIME_CHANNELS.event,
      payload: {
        documentId: 'doc-1',
        event: { type: 'error', code: 'codex_turn_start_failed', message: 'Turn failed.' },
      },
    })
    f.client.startTurnError = undefined
    await expect(f.runtime.startTurn(f.owner, 'doc-1', 'A fresh explicit turn')).resolves.toEqual({
      turnId: 'turn-1',
    })
    expect(f.calls.filter((call) => call === 'client.startTurn')).toHaveLength(2)
  })

  it('serializes duplicate starts and a close racing document startup', async () => {
    const f = harness({ deferManagerStart: true })
    const first = f.runtime.registerDocument({
      documentId: 'doc-1',
      owner: f.owner,
      registration: registration(),
    })
    await flush()
    await expect(
      f.runtime.registerDocument({
        documentId: 'doc-1',
        owner: f.owner,
        registration: registration(),
      }),
    ).rejects.toMatchObject({ code: 'codex_document_exists' })
    const closing = f.runtime.closeDocument('doc-1')
    f.managerStart.resolve(f.client as unknown as CodexAppServerClient)

    await expect(first).rejects.toMatchObject({ code: 'codex_session_closed' })
    await expect(closing).resolves.toBeUndefined()
    expect(f.manager.stop).toHaveBeenCalledOnce()
    expect(f.bridge.close).toHaveBeenCalledOnce()
  })

  it('quarantines new document starts while logout drains an in-flight startup', async () => {
    const f = harness({ deferManagerStart: true })
    const starting = f.runtime.registerDocument({
      documentId: 'doc-1',
      owner: f.owner,
      registration: registration(),
    })
    await flush()
    const loggingOut = f.runtime.logout()

    await expect(
      f.runtime.registerDocument({
        documentId: 'doc-2',
        owner: new FakeOwner(),
        registration: registration(),
      }),
    ).rejects.toMatchObject({ code: 'codex_runtime_busy' })

    f.managerStart.resolve(f.client as unknown as CodexAppServerClient)
    await expect(starting).rejects.toMatchObject({ code: 'codex_session_closed' })
    await expect(loggingOut).resolves.toBeUndefined()
    expect(f.manager.stop).toHaveBeenCalledOnce()
  })

  it('rejects cross-document owners and concurrent turns', async () => {
    const f = harness()
    const other = new FakeOwner()
    await startDocument(f)
    await f.runtime.startTurn(f.owner, 'doc-1', 'First')

    await expect(f.runtime.startTurn(other, 'doc-1', 'Cross document')).rejects.toMatchObject({
      code: 'codex_document_unavailable',
    })
    await expect(f.runtime.startTurn(f.owner, 'doc-1', 'Duplicate')).rejects.toMatchObject({
      code: 'codex_turn_in_progress',
    })
  })

  it('reserves a turn before an asynchronous auth refresh can admit a concurrent start', async () => {
    const f = harness()
    await startDocument(f)
    const auth = deferred<{ loggedIn: boolean }>()
    f.authClient.getValidAccountStatus.mockImplementation(() => auth.promise)

    const first = f.runtime.startTurn(f.owner, 'doc-1', 'First')
    const second = expect(f.runtime.startTurn(f.owner, 'doc-1', 'Duplicate')).rejects.toMatchObject(
      { code: 'codex_turn_in_progress' },
    )
    await flush()
    expect(f.calls.filter((call) => call === 'client.startTurn')).toHaveLength(0)

    auth.resolve({ loggedIn: true })
    await expect(first).resolves.toEqual({ turnId: 'turn-1' })
    await second
    expect(f.calls.filter((call) => call === 'client.startTurn')).toHaveLength(1)
  })

  it('blocks fresh turns while close is quiesced and resumes the same document after cancellation', async () => {
    const f = harness()
    await startDocument(f)

    await f.runtime.quiesceDocument('doc-1')
    await expect(f.runtime.startTurn(f.owner, 'doc-1', 'Blocked')).rejects.toMatchObject({
      code: 'codex_document_unavailable',
    })

    f.runtime.resumeDocument('doc-1')
    await expect(f.runtime.startTurn(f.owner, 'doc-1', 'Allowed')).resolves.toEqual({
      turnId: 'turn-1',
    })
  })

  it('resumes runtime and renderer tools when a tab close is cancelled', async () => {
    const f = harness()
    await startDocument(f)
    const tools = {
      quiesceDocument: vi.fn(async () => undefined),
      resumeDocument: vi.fn(),
      quiesceSessions: vi.fn(async () => undefined),
      resumeSessions: vi.fn(),
      closeDocument: vi.fn(),
      closeSessions: vi.fn(),
    }

    const close = await prepareCodexDocumentClose(f.runtime, tools, 'doc-1')
    close.rollback()

    expect(tools.resumeDocument).toHaveBeenCalledWith('doc-1')
    await expect(f.runtime.startTurn(f.owner, 'doc-1', 'Still usable')).resolves.toEqual({
      turnId: 'turn-1',
    })
    expect(tools.closeDocument).not.toHaveBeenCalled()
  })

  it('rolls back runtime quiescence when interrupt acquisition rejects', async () => {
    const f = harness()
    await startDocument(f)
    await f.runtime.startTurn(f.owner, 'doc-1', 'Running')
    f.client.interruptTurnError = new Error('untrusted raw interrupt failure')
    const tools = {
      quiesceDocument: vi.fn(async () => undefined),
      resumeDocument: vi.fn(),
      quiesceSessions: vi.fn(async () => undefined),
      resumeSessions: vi.fn(),
      closeDocument: vi.fn(),
      closeSessions: vi.fn(),
    }

    await expect(prepareCodexDocumentClose(f.runtime, tools, 'doc-1')).rejects.toMatchObject({
      code: 'codex_turn_interrupt_failed',
    })
    expect(tools.resumeDocument).not.toHaveBeenCalled()

    f.client.emit({
      method: 'turn/completed',
      params: { threadId: 'thread-1', turn: { id: 'turn-1', status: 'interrupted' } },
    })
    f.client.interruptTurnError = undefined
    await expect(f.runtime.startTurn(f.owner, 'doc-1', 'Still usable')).resolves.toEqual({
      turnId: 'turn-1',
    })
  })

  it('fully settles multi-document quiescence before rolling back a partial rejection', async () => {
    const f = harness()
    await startDocument(f)
    await f.runtime.registerDocument({
      documentId: 'doc-2',
      owner: f.owner,
      registration: registration(),
    })
    const late = deferred<void>()
    const originalQuiesce = f.runtime.quiesceDocument.bind(f.runtime)
    const quiesce = vi
      .spyOn(f.runtime, 'quiesceDocument')
      .mockImplementation(async (documentId) => {
        if (documentId === 'doc-1') throw new ShellCodexRuntimeError('codex_turn_interrupt_failed')
        await late.promise
        await originalQuiesce(documentId)
      })
    const tools = {
      quiesceDocument: vi.fn(async () => undefined),
      resumeDocument: vi.fn(),
      quiesceSessions: vi.fn(async () => undefined),
      resumeSessions: vi.fn(),
      closeDocument: vi.fn(),
      closeSessions: vi.fn(),
    }

    let rejected = false
    const closing = expect(prepareCodexDocumentsClose(f.runtime, tools))
      .rejects.toMatchObject({ code: 'codex_turn_interrupt_failed' })
      .finally(() => {
        rejected = true
      })
    await flush()
    expect(rejected).toBe(false)
    late.resolve()
    await closing
    quiesce.mockRestore()

    expect(tools.resumeSessions).not.toHaveBeenCalled()
    await expect(f.runtime.startTurn(f.owner, 'doc-2', 'Resumed')).resolves.toEqual({
      turnId: 'turn-1',
    })
  })

  it('gates and drains a document registration that is still starting during close preparation', async () => {
    const f = harness({ deferManagerStart: true })
    const tools = {
      quiesceDocument: vi.fn(async () => undefined),
      resumeDocument: vi.fn(),
      quiesceSessions: vi.fn(async () => undefined),
      resumeSessions: vi.fn(),
      closeDocument: vi.fn(),
      closeSessions: vi.fn(),
    }
    const starting = f.runtime.registerDocument({
      documentId: 'doc-1',
      owner: f.owner,
      registration: registration(),
    })
    await flush()
    let prepared = false
    const close = prepareCodexDocumentClose(f.runtime, tools, 'doc-1').then((value) => {
      prepared = true
      return value
    })
    await flush()
    expect(prepared).toBe(false)
    await expect(
      f.runtime.registerDocument({
        documentId: 'doc-1',
        owner: f.owner,
        registration: registration(),
      }),
    ).rejects.toMatchObject({ code: 'codex_runtime_busy' })

    f.managerStart.resolve(f.client as unknown as CodexAppServerClient)
    await expect(starting).rejects.toMatchObject({ code: 'codex_session_closed' })
    const preparation = await close
    preparation.rollback()

    await expect(
      f.runtime.registerDocument({
        documentId: 'doc-1',
        owner: f.owner,
        registration: registration(),
      }),
    ).resolves.toBeDefined()
    const laterWindowClose = await prepareCodexDocumentsClose(f.runtime, tools)
    laterWindowClose.rollback()
  })

  it('retains an approved document-close tombstone so no runtime can reactivate before detach', async () => {
    const f = harness()
    await startDocument(f)
    const tools = {
      quiesceDocument: vi.fn(async () => undefined),
      resumeDocument: vi.fn(),
      quiesceSessions: vi.fn(async () => undefined),
      resumeSessions: vi.fn(),
      closeDocument: vi.fn(),
      closeSessions: vi.fn(),
    }

    const preparation = await prepareCodexDocumentClose(f.runtime, tools, 'doc-1')
    await preparation.commit()

    expect(tools.closeDocument).toHaveBeenCalledWith('doc-1')
    await expect(
      f.runtime.registerDocument({
        documentId: 'doc-1',
        owner: f.owner,
        registration: registration(),
      }),
    ).rejects.toMatchObject({ code: 'codex_runtime_busy' })

    preparation.finalize()
    await expect(
      f.runtime.registerDocument({
        documentId: 'doc-1',
        owner: f.owner,
        registration: registration(),
      }),
    ).resolves.toBeDefined()
    const laterGlobalClose = await prepareCodexDocumentsClose(f.runtime, tools)
    laterGlobalClose.rollback()
  })

  it('does not let a losing overlapping document close resume the winning preparation', async () => {
    const f = harness()
    await startDocument(f)
    const late = deferred<void>()
    const tools = {
      quiesceDocument: vi.fn(async () => late.promise),
      resumeDocument: vi.fn(),
      quiesceSessions: vi.fn(async () => undefined),
      resumeSessions: vi.fn(),
      closeDocument: vi.fn(),
      closeSessions: vi.fn(),
    }
    const firstPending = prepareCodexDocumentClose(f.runtime, tools, 'doc-1')
    await flush()

    await expect(prepareCodexDocumentClose(f.runtime, tools, 'doc-1')).rejects.toMatchObject({
      code: 'codex_runtime_busy',
    })
    await expect(f.runtime.startTurn(f.owner, 'doc-1', 'Must stay blocked')).rejects.toMatchObject({
      code: 'codex_document_unavailable',
    })
    expect(tools.resumeDocument).not.toHaveBeenCalled()

    late.resolve()
    const first = await firstPending
    first.rollback()
    expect(tools.resumeDocument).toHaveBeenCalledTimes(1)
    await expect(f.runtime.startTurn(f.owner, 'doc-1', 'Resumed once')).resolves.toEqual({
      turnId: 'turn-1',
    })
  })

  it('globally drains a registration racing window close and allows retry only after cancellation', async () => {
    const f = harness({ deferManagerStart: true })
    const tools = {
      quiesceDocument: vi.fn(async () => undefined),
      resumeDocument: vi.fn(),
      quiesceSessions: vi.fn(async () => undefined),
      resumeSessions: vi.fn(),
      closeDocument: vi.fn(),
      closeSessions: vi.fn(),
    }
    const starting = f.runtime.registerDocument({
      documentId: 'doc-1',
      owner: f.owner,
      registration: registration(),
    })
    await flush()
    let prepared = false
    const closing = prepareCodexDocumentsClose(f.runtime, tools).then((value) => {
      prepared = true
      return value
    })
    await flush()
    expect(prepared).toBe(false)

    f.managerStart.resolve(f.client as unknown as CodexAppServerClient)
    await expect(starting).rejects.toMatchObject({ code: 'codex_session_closed' })
    const preparation = await closing
    preparation.rollback()

    await expect(
      f.runtime.registerDocument({
        documentId: 'doc-2',
        owner: f.owner,
        registration: registration(),
      }),
    ).resolves.toBeDefined()
  })

  it('does not let a losing global close resume sessions owned by the first close', async () => {
    const f = harness()
    await startDocument(f)
    const tools = {
      quiesceDocument: vi.fn(async () => undefined),
      resumeDocument: vi.fn(),
      quiesceSessions: vi.fn(async () => undefined),
      resumeSessions: vi.fn(),
      closeDocument: vi.fn(),
      closeSessions: vi.fn(),
    }
    const first = await prepareCodexDocumentsClose(f.runtime, tools)

    await expect(prepareCodexDocumentsClose(f.runtime, tools)).rejects.toMatchObject({
      code: 'codex_runtime_busy',
    })
    await expect(f.runtime.startTurn(f.owner, 'doc-1', 'Must stay blocked')).rejects.toMatchObject({
      code: 'codex_document_unavailable',
    })
    expect(tools.resumeSessions).not.toHaveBeenCalled()

    first.rollback()
    expect(tools.resumeSessions).toHaveBeenCalledTimes(1)
    await expect(f.runtime.startTurn(f.owner, 'doc-1', 'Resumed once')).resolves.toEqual({
      turnId: 'turn-1',
    })
  })

  it('releases a committed window tombstone only after identity-removal finalization', async () => {
    const f = harness()
    await startDocument(f)
    const tools = {
      quiesceDocument: vi.fn(async () => undefined),
      resumeDocument: vi.fn(),
      quiesceSessions: vi.fn(async () => undefined),
      resumeSessions: vi.fn(),
      closeDocument: vi.fn(),
      closeSessions: vi.fn(),
    }
    const preparation = await prepareCodexDocumentsClose(f.runtime, tools)
    await preparation.commit()

    await expect(
      f.runtime.registerDocument({
        documentId: 'doc-1',
        owner: f.owner,
        registration: registration(),
      }),
    ).rejects.toMatchObject({ code: 'codex_runtime_busy' })

    preparation.finalize()
    await expect(
      f.runtime.registerDocument({
        documentId: 'doc-1',
        owner: f.owner,
        registration: registration(),
      }),
    ).resolves.toBeDefined()
  })

  it('resumes every runtime and renderer tool session when window close is cancelled', async () => {
    const f = harness()
    await startDocument(f)
    const tools = {
      quiesceDocument: vi.fn(async () => undefined),
      resumeDocument: vi.fn(),
      quiesceSessions: vi.fn(async () => undefined),
      resumeSessions: vi.fn(),
      closeDocument: vi.fn(),
      closeSessions: vi.fn(),
    }

    const close = await prepareCodexDocumentsClose(f.runtime, tools)
    await expect(
      f.runtime.registerDocument({
        documentId: 'doc-2',
        owner: f.owner,
        registration: registration(),
      }),
    ).rejects.toMatchObject({ code: 'codex_runtime_busy' })
    close.rollback()

    expect(tools.resumeSessions).toHaveBeenCalledOnce()
    await expect(f.runtime.startTurn(f.owner, 'doc-1', 'Still usable')).resolves.toEqual({
      turnId: 'turn-1',
    })
    expect(tools.closeSessions).not.toHaveBeenCalled()
  })

  it('holds global registration gates until credential logout finishes', async () => {
    const f = harness()
    await startDocument(f)
    const authority = deferred<void>()
    const tools = {
      quiesceDocument: vi.fn(async () => undefined),
      resumeDocument: vi.fn(),
      quiesceSessions: vi.fn(async () => undefined),
      resumeSessions: vi.fn(),
      closeDocument: vi.fn(),
      closeSessions: vi.fn(),
    }
    const logout = logoutWithCodexClose(f.runtime, tools, () => authority.promise)
    await flush()

    await expect(
      f.runtime.registerDocument({
        documentId: 'doc-2',
        owner: f.owner,
        registration: registration(),
      }),
    ).rejects.toMatchObject({ code: 'codex_runtime_busy' })

    f.authClient.getValidAccountStatus.mockResolvedValue({ loggedIn: false })
    authority.resolve()
    await expect(logout).resolves.toBeUndefined()
    await expect(
      f.runtime.registerDocument({
        documentId: 'doc-2',
        owner: f.owner,
        registration: registration(),
      }),
    ).rejects.toMatchObject({ code: 'auth_required' })
  })

  it('keeps cleanup closed through credential rejection, then releases the retry lease', async () => {
    const f = harness()
    await startDocument(f)
    const authority = deferred<void>()
    const tools = {
      quiesceDocument: vi.fn(async () => undefined),
      resumeDocument: vi.fn(),
      quiesceSessions: vi.fn(async () => undefined),
      resumeSessions: vi.fn(),
      closeDocument: vi.fn(),
      closeSessions: vi.fn(),
    }
    const logout = expect(
      logoutWithCodexClose(f.runtime, tools, () => authority.promise),
    ).rejects.toMatchObject({ code: 'auth_logout_failed' })
    await flush()
    await expect(
      f.runtime.registerDocument({
        documentId: 'doc-2',
        owner: f.owner,
        registration: registration(),
      }),
    ).rejects.toMatchObject({ code: 'codex_runtime_busy' })
    authority.reject(new Error('raw credential failure'))
    await logout

    expect(f.manager.stop).toHaveBeenCalledOnce()
    await expect(
      f.runtime.registerDocument({
        documentId: 'doc-2',
        owner: f.owner,
        registration: registration(),
      }),
    ).resolves.toBeDefined()
  })
})

describe('prepared window close', () => {
  it('rolls back when the first close preparation returns false', async () => {
    const preparation = { commit: vi.fn(async () => undefined), rollback: vi.fn() }

    await expect(runCodexPreparedClose(preparation, [vi.fn(async () => false)])).resolves.toBe(
      false,
    )

    expect(preparation.commit).not.toHaveBeenCalled()
    expect(preparation.rollback).toHaveBeenCalledOnce()
  })

  it('rolls back when a close preparation throws', async () => {
    const preparation = { commit: vi.fn(async () => undefined), rollback: vi.fn() }

    await expect(
      runCodexPreparedClose(preparation, [async () => Promise.reject(new Error('latex_failed'))]),
    ).rejects.toThrow('latex_failed')

    expect(preparation.commit).not.toHaveBeenCalled()
    expect(preparation.rollback).toHaveBeenCalledOnce()
  })

  it('stops after a partial multi-tab refusal and rolls back once', async () => {
    const preparation = { commit: vi.fn(async () => undefined), rollback: vi.fn() }
    const first = vi.fn(async () => true)
    const second = vi.fn(async () => false)
    const third = vi.fn(async () => true)

    await expect(runCodexPreparedClose(preparation, [first, second, third])).resolves.toBe(false)

    expect(first).toHaveBeenCalledOnce()
    expect(second).toHaveBeenCalledOnce()
    expect(third).not.toHaveBeenCalled()
    expect(preparation.rollback).toHaveBeenCalledOnce()
  })
})

describe('WisUsage bridge fetch', () => {
  it('uses AuthClient.fetchWithAuth and preserves the fixed Codex model', async () => {
    const upstream = vi.fn(async () => new Response(null, { status: 200 }))
    const fetchWithAuth = vi.fn(async (request: (accessToken: string) => Promise<Response>) =>
      request('private-access-token'),
    )
    const fetchBridge = createWisUsageBridgeFetch({ fetchWithAuth }, upstream)
    const request = {
      model: 'openai/gpt-5.6-sol',
      max_tokens: 1024,
      messages: [{ role: 'user', content: 'hello' }],
    } as MessagesRequest

    await fetchBridge(request, new AbortController().signal)

    expect(fetchWithAuth).toHaveBeenCalledOnce()
    expect(upstream).toHaveBeenCalledOnce()
    const [url, init] = upstream.mock.calls[0]!
    expect(url).toBe('https://wisusage.dev.atominnolab.com/v1/messages')
    expect(init?.headers).toEqual({
      Authorization: 'Bearer private-access-token',
      'Content-Type': 'application/json',
      'x-req-location': 'sg',
    })
    expect(JSON.parse(String(init?.body))).toMatchObject({ model: 'openai/gpt-5.6-sol' })
  })

  it('fails closed before auth when a different model reaches the Shell boundary', async () => {
    const fetchWithAuth = vi.fn()
    const fetchBridge = createWisUsageBridgeFetch({ fetchWithAuth }, vi.fn())

    await expect(
      fetchBridge(
        { model: 'override/model', messages: [], max_tokens: 1 } as unknown as MessagesRequest,
        new AbortController().signal,
      ),
    ).rejects.toMatchObject({ code: 'codex_model_mismatch' })
    expect(fetchWithAuth).not.toHaveBeenCalled()
  })
})

describe('Codex app quit coordination', () => {
  it('prevents quit until one bounded runtime shutdown settles, then retries once', async () => {
    const shutdown = deferred<void>()
    const quit = vi.fn()
    const diagnostics = vi.fn()
    const handler = createCodexBeforeQuitHandler({
      shutdown: vi.fn(() => shutdown.promise),
      quit,
      diagnostics,
    })
    const first = { preventDefault: vi.fn() }
    const duplicate = { preventDefault: vi.fn() }

    handler(first)
    handler(duplicate)
    expect(first.preventDefault).toHaveBeenCalledOnce()
    expect(duplicate.preventDefault).toHaveBeenCalledOnce()
    expect(quit).not.toHaveBeenCalled()
    shutdown.resolve()
    await flush()
    expect(quit).toHaveBeenCalledOnce()

    const retried = { preventDefault: vi.fn() }
    handler(retried)
    expect(retried.preventDefault).not.toHaveBeenCalled()
    expect(diagnostics).not.toHaveBeenCalled()
  })
})
