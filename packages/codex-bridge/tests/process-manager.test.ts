import { EventEmitter } from 'node:events'
import { tmpdir } from 'node:os'
import { join } from 'node:path'
import { PassThrough, Writable } from 'node:stream'
import { describe, expect, it, vi } from 'vitest'
import {
  CodexProcessManager,
  type ChildProcessAdapter,
  type CodexSpawn,
  type OwnedCodexDirectories,
} from '../src/process-manager.js'

class FakeChild extends EventEmitter implements ChildProcessAdapter {
  readonly stdin: Writable
  readonly stdout = new PassThrough()
  readonly stderr = new PassThrough()
  readonly kills: Array<NodeJS.Signals> = []
  exitCode: number | null = null
  signalCode: NodeJS.Signals | null = null
  killed = false
  errorOnKill = false

  constructor(stdin: Writable = new PassThrough()) {
    super()
    this.stdin = stdin
  }

  kill(signal: NodeJS.Signals = 'SIGTERM'): boolean {
    this.kills.push(signal)
    this.killed = true
    if (this.errorOnKill) this.emit('error', new Error('private kill failure'))
    return true
  }

  spawned(): void {
    this.emit('spawn')
  }

  exited(code: number | null = 0, signal: NodeJS.Signals | null = null): void {
    this.exitCode = code
    this.signalCode = signal
    this.emit('exit', code, signal)
  }
}

function createFixture(
  fixtureOptions: {
    autoVersion?: boolean
    autoServerSpawn?: boolean
    serverStdin?: Writable
    removeDirectories?: (directories: OwnedCodexDirectories) => Promise<void>
    mcp?: { url: string; secret: string }
    executablePath?: string
    versionOutput?: string
  } = {},
) {
  const version = new FakeChild()
  const server = new FakeChild(fixtureOptions.serverStdin)
  const children = [version, server]
  const calls: Array<{
    executable: string
    args: readonly string[]
    env: NodeJS.ProcessEnv
    cwd: string
  }> = []
  const spawn: CodexSpawn = (executable, args, options) => {
    calls.push({ executable, args, env: options.env, cwd: options.cwd })
    const child = children.shift()
    if (!child) throw new Error('unexpected spawn')
    if (child === version || fixtureOptions.autoServerSpawn !== false)
      queueMicrotask(() => child.spawned())
    if (child === version && fixtureOptions.autoVersion !== false) {
      queueMicrotask(() => {
        child.stdout.end(fixtureOptions.versionOutput ?? 'codex-cli 0.147.0\n')
        child.exited(0)
      })
    }
    return child
  }
  const root = join(tmpdir(), 'wiswork-codex-unit-1')
  const directories: OwnedCodexDirectories = {
    root,
    codexHome: join(root, 'home'),
    cwd: join(root, 'workspace'),
  }
  const createDirectories = vi.fn(async () => directories)
  const removeDirectories = vi.fn(fixtureOptions.removeDirectories ?? (async () => undefined))
  const diagnostics = vi.fn()
  const manager = new CodexProcessManager({
    executablePath: fixtureOptions.executablePath ?? '/opt/wiswork/codex',
    bridge: { baseUrl: 'http://127.0.0.1:43123', secret: 'bridge-secret-private' },
    mcp: fixtureOptions.mcp,
    developerInstructions: 'Fixed host policy.',
    spawn,
    createDirectories,
    removeDirectories,
    diagnostics,
    startupTimeoutMs: 100,
    gracefulShutdownMs: 10,
    terminateTimeoutMs: 10,
    killTimeoutMs: 10,
  })
  return {
    manager,
    version,
    server,
    calls,
    directories,
    createDirectories,
    removeDirectories,
    diagnostics,
  }
}

async function tick(): Promise<void> {
  await new Promise((resolve) => setImmediate(resolve))
}

describe('pinned Codex app-server process manager', () => {
  it('requires an absolute executable and a strict loopback bridge URL', () => {
    const base = {
      bridge: { baseUrl: 'http://127.0.0.1:1234', secret: 'secret' },
      developerInstructions: 'policy',
    }
    expect(() => new CodexProcessManager({ ...base, executablePath: 'codex' })).toThrow(
      'codex_executable_must_be_absolute',
    )
    expect(
      () =>
        new CodexProcessManager({
          ...base,
          executablePath: '/codex',
          bridge: { baseUrl: 'https://example.com', secret: 'secret' },
        }),
    ).toThrow('invalid_codex_bridge')
    expect(
      () => new CodexProcessManager({ ...base, executablePath: '/opt/wiswork/helper' }),
    ).toThrow('codex_executable_name_invalid')
    expect(
      () =>
        new CodexProcessManager({
          ...base,
          executablePath: '/codex',
          createDirectories: async () => ({ root: '/', codexHome: '/home', cwd: '/workspace' }),
        }),
    ).toThrow('codex_temp_adapters_must_be_paired')
    expect(
      () =>
        new CodexProcessManager({
          ...base,
          executablePath: '/codex',
          developerInstructions: 'x'.repeat(65_537),
        }),
    ).toThrow('invalid_developer_instructions')
    expect(
      () =>
        new CodexProcessManager({
          ...base,
          executablePath: '/codex',
          bridge: { baseUrl: 'http://127.0.0.1:1234', secret: 'x'.repeat(4_097) },
        }),
    ).toThrow('invalid_codex_bridge')
  })

  it('refuses broad or unowned cleanup roots before spawning', async () => {
    const removeDirectories = vi.fn(async () => undefined)
    const manager = new CodexProcessManager({
      executablePath: '/opt/wiswork/codex',
      bridge: { baseUrl: 'http://127.0.0.1:1234', secret: 'secret' },
      developerInstructions: 'policy',
      createDirectories: async () => ({ root: '/', codexHome: '/home', cwd: '/workspace' }),
      removeDirectories,
    })

    await expect(manager.start()).rejects.toMatchObject({ code: 'invalid_owned_directories' })
    expect(removeDirectories).not.toHaveBeenCalled()
  })

  it('verifies the exact pinned version before spawning with strict config', async () => {
    const fixture = createFixture()
    const starting = fixture.manager.start()
    await tick()
    const client = await starting

    expect(fixture.calls[0]).toMatchObject({
      executable: '/opt/wiswork/codex',
      args: ['--version'],
      cwd: fixture.directories.cwd,
    })
    expect(fixture.calls[1]?.executable).toBe('/opt/wiswork/codex')
    expect(fixture.calls[1]?.args).toEqual([
      'app-server',
      '--strict-config',
      '--stdio',
      '-c',
      'model_provider="wiswork"',
      '-c',
      'model_providers.wiswork.name="WisWork"',
      '-c',
      'model_providers.wiswork.base_url="http://127.0.0.1:43123/v1"',
      '-c',
      'model_providers.wiswork.env_key="WISWORK_CODEX_TOKEN"',
      '-c',
      'model_providers.wiswork.wire_api="responses"',
      '-c',
      'features.shell_tool=false',
      '-c',
      'features.unified_exec=false',
      '-c',
      'features.code_mode=true',
      '-c',
      'features.code_mode_host=true',
      '-c',
      'tools.update_plan.enabled=false',
      '-c',
      'features.multi_agent=false',
    ])
    expect(fixture.calls[1]?.cwd).toBe(fixture.directories.cwd)
    expect(fixture.calls[1]?.args.join(' ')).not.toContain('bridge-secret-private')
    expect(fixture.calls[1]?.env).toMatchObject({
      CODEX_HOME: fixture.directories.codexHome,
      WISWORK_CODEX_TOKEN: 'bridge-secret-private',
      CODEX_CODE_MODE_HOST_PATH:
        process.platform === 'win32'
          ? '\\opt\\wiswork\\codex-code-mode-host.exe'
          : '/opt/wiswork/codex-code-mode-host',
      NO_PROXY: '127.0.0.1,localhost',
      no_proxy: '127.0.0.1,localhost',
    })
    expect(fixture.calls[1]?.env).not.toHaveProperty('HTTP_PROXY')
    expect(fixture.calls[1]?.env).not.toHaveProperty('WISPAPER_ACCESS_TOKEN')

    fixture.server.exited(0)
    await fixture.manager.stop()
    expect(client).toBeDefined()
  })

  it('starts the pinned app-server-only component without a duplicate subcommand', async () => {
    const executablePath = `/opt/wiswork/codex-app-server${process.platform === 'win32' ? '.exe' : ''}`
    const fixture = createFixture({
      executablePath,
      versionOutput: 'codex-app-server 0.147.0\n',
    })
    const starting = fixture.manager.start()
    await tick()
    await starting

    expect(fixture.calls[0]).toMatchObject({ executable: executablePath, args: ['--version'] })
    expect(fixture.calls[1]?.executable).toBe(executablePath)
    expect(fixture.calls[1]?.args[0]).toBe('--strict-config')
    expect(fixture.calls[1]?.args).not.toContain('app-server')

    fixture.server.exited(0)
    await fixture.manager.stop()
  })

  it('configures the exact Codex 0.147 HTTP MCP keys with its secret only in child env', async () => {
    const fixture = createFixture({
      mcp: {
        url: `http://127.0.0.1:44321/mcp/${Buffer.alloc(32, 7).toString('base64url')}`,
        secret: Buffer.alloc(32, 8).toString('base64url'),
      },
    })
    const started = fixture.manager.start()
    await tick()
    await started

    expect(fixture.calls[1]?.args).toContain(
      `mcp_servers.wiswork.url="http://127.0.0.1:44321/mcp/${Buffer.alloc(32, 7).toString('base64url')}"`,
    )
    expect(fixture.calls[1]?.args).toContain(
      'mcp_servers.wiswork.bearer_token_env_var="WISWORK_MCP_TOKEN"',
    )
    expect(fixture.calls[1]?.args.join(' ')).not.toContain(
      Buffer.alloc(32, 8).toString('base64url'),
    )
    expect(fixture.calls[1]?.env.WISWORK_MCP_TOKEN).toBe(Buffer.alloc(32, 8).toString('base64url'))
    fixture.server.exited(0)
    await fixture.manager.stop()
  })

  it('rejects non-loopback or unscoped MCP endpoints before creating temp state', () => {
    const base = {
      executablePath: '/opt/wiswork/codex',
      bridge: { baseUrl: 'http://127.0.0.1:1234', secret: 'secret' },
      developerInstructions: 'policy',
    }
    for (const url of [
      'https://example.com/mcp/session',
      'http://localhost:1234/mcp/session',
      'http://127.0.0.1:1234/mcp',
      'http://127.0.0.1:1234/mcp/session?document=private',
    ]) {
      expect(
        () =>
          new CodexProcessManager({
            ...base,
            mcp: { url, secret: Buffer.alloc(32, 9).toString('base64url') },
          }),
      ).toThrow('invalid_codex_mcp')
    }
    expect(
      () =>
        new CodexProcessManager({
          ...base,
          mcp: {
            url: `http://127.0.0.1:1234/mcp/${Buffer.alloc(32, 9).toString('base64url')}`,
            secret: 'short-secret',
          },
        }),
    ).toThrow('invalid_codex_mcp')
  })

  it('shares a concurrent start but rejects a duplicate after startup', async () => {
    const fixture = createFixture()
    const first = fixture.manager.start()
    const concurrent = fixture.manager.start()
    expect(concurrent).toBe(first)
    await tick()
    await first
    await expect(fixture.manager.start()).rejects.toMatchObject({
      code: 'codex_process_already_started',
    })
    fixture.server.exited(0)
    await fixture.manager.stop()
  })

  it('fails closed on version mismatch and removes owned directories', async () => {
    const fixture = createFixture({ autoVersion: false })
    const start = fixture.manager.start()
    await tick()
    fixture.version.stdout.end('codex-cli 0.148.0 private\n')
    fixture.version.exited(0)

    await expect(start).rejects.toMatchObject({ code: 'codex_version_mismatch' })
    expect(fixture.calls).toHaveLength(1)
    expect(fixture.removeDirectories).toHaveBeenCalledWith(fixture.directories)
    expect(JSON.stringify(fixture.diagnostics.mock.calls)).not.toContain('private')
  })

  it('bounds version output and its verification deadline', async () => {
    const output = createFixture({ autoVersion: false })
    const tooLarge = output.manager.start()
    await tick()
    output.version.stdout.write('x'.repeat(65_537))
    output.version.exited(null, 'SIGKILL')
    await expect(tooLarge).rejects.toMatchObject({ code: 'codex_version_output_limit' })
    expect(output.version.kills).toContain('SIGKILL')
    expect(output.removeDirectories).toHaveBeenCalledOnce()

    vi.useFakeTimers()
    try {
      const timeout = createFixture({ autoVersion: false })
      const pending = timeout.manager.start()
      const rejection = expect(pending).rejects.toMatchObject({
        code: 'codex_version_timeout',
      })
      await vi.advanceTimersByTimeAsync(100)
      timeout.version.exited(null, 'SIGKILL')
      await rejection
      expect(timeout.version.kills).toContain('SIGKILL')
      expect(timeout.removeDirectories).toHaveBeenCalledOnce()
    } finally {
      vi.useRealTimers()
    }
  })

  it('rejects start on early child error or exit and cleans up', async () => {
    const errored = createFixture({ autoServerSpawn: false })
    const errorStart = errored.manager.start()
    await tick()
    errored.server.emit('error', new Error('private spawn failure'))
    errored.server.exited(null, 'SIGKILL')
    await expect(errorStart).rejects.toMatchObject({ code: 'codex_process_start_failed' })
    expect(errored.removeDirectories).toHaveBeenCalledOnce()

    const exited = createFixture({ autoServerSpawn: false })
    const exitStart = exited.manager.start()
    await tick()
    exited.server.exited(9)
    await expect(exitStart).rejects.toMatchObject({ code: 'codex_process_exited' })
    expect(exited.removeDirectories).toHaveBeenCalledOnce()
  })

  it('rejects every pending RPC request when the child crashes', async () => {
    const fixture = createFixture()
    const client = fixture.manager.start()
    await tick()
    const app = await client
    const initialization = app.initialize()
    fixture.server.exited(2)

    await expect(initialization).rejects.toMatchObject({ code: 'codex_process_exited' })
    await expect(fixture.manager.crashed).resolves.toMatchObject({
      code: 'codex_process_exited',
    })
    expect(fixture.removeDirectories).toHaveBeenCalledOnce()
  })

  it('resolves crash reporting even when owned-directory cleanup fails', async () => {
    const fixture = createFixture({
      removeDirectories: async () => {
        throw new Error('private cleanup failure')
      },
    })
    const app = await fixture.manager.start()
    const initialization = app.initialize()
    fixture.server.exited(2)

    await expect(initialization).rejects.toMatchObject({ code: 'codex_process_exited' })
    await expect(fixture.manager.crashed).resolves.toMatchObject({
      code: 'codex_process_exited',
    })
    expect(fixture.diagnostics).toHaveBeenCalledWith({ code: 'codex_temp_cleanup_failed' })
    expect(fixture.diagnostics).toHaveBeenCalledWith({ code: 'codex_process_exited' })
  })

  it('does not clean temp state until a killed verifier confirms exit', async () => {
    vi.useFakeTimers()
    try {
      const fixture = createFixture({ autoVersion: false })
      fixture.version.errorOnKill = true
      const start = fixture.manager.start()
      const rejection = expect(start).rejects.toMatchObject({
        code: 'codex_process_termination_timeout',
      })
      await vi.advanceTimersByTimeAsync(100)
      expect(fixture.version.kills).toEqual(['SIGKILL'])
      expect(fixture.removeDirectories).not.toHaveBeenCalled()

      await vi.advanceTimersByTimeAsync(10)
      await rejection
      expect(fixture.removeDirectories).not.toHaveBeenCalled()
    } finally {
      vi.useRealTimers()
    }
  })

  it('drains stderr but exposes only bounded diagnostic codes', async () => {
    const fixture = createFixture()
    const start = fixture.manager.start()
    await tick()
    await start
    fixture.server.stderr.write(`private prompt bridge-secret-private ${'x'.repeat(70_000)}`)
    await tick()

    expect(fixture.diagnostics).toHaveBeenCalledWith({ code: 'codex_stderr_output' })
    expect(fixture.diagnostics).toHaveBeenCalledWith({ code: 'codex_stderr_truncated' })
    expect(fixture.diagnostics.mock.calls).toHaveLength(2)
    expect(JSON.stringify(fixture.diagnostics.mock.calls)).not.toContain('private')
    fixture.server.exited(0)
    await fixture.manager.stop()
  })

  it('ends stdin, escalates to SIGTERM then SIGKILL, and stops idempotently', async () => {
    vi.useFakeTimers()
    try {
      const fixture = createFixture()
      const started = fixture.manager.start()
      await vi.runOnlyPendingTimersAsync()
      await started
      const first = fixture.manager.stop()
      const second = fixture.manager.stop()
      expect(second).toBe(first)
      expect(fixture.server.stdin.writableEnded).toBe(true)

      await vi.advanceTimersByTimeAsync(10)
      expect(fixture.server.kills).toEqual(['SIGTERM'])
      await vi.advanceTimersByTimeAsync(10)
      expect(fixture.server.kills).toEqual(['SIGTERM', 'SIGKILL'])
      fixture.server.exited(null, 'SIGKILL')
      await first
      expect(fixture.removeDirectories).toHaveBeenCalledOnce()
      await expect(fixture.manager.stop()).resolves.toBeUndefined()
    } finally {
      vi.useRealTimers()
    }
  })

  it('does not let a stalled stdin flush bypass bounded termination', async () => {
    vi.useFakeTimers()
    try {
      const stalled = new Writable({
        write(_chunk, _encoding, _callback) {
          // Simulate a child that stopped reading without closing its pipe.
        },
      })
      const fixture = createFixture({ serverStdin: stalled })
      const started = fixture.manager.start()
      await vi.runOnlyPendingTimersAsync()
      const app = await started
      const initializing = app.initialize()
      void initializing.catch(() => undefined)
      const stopping = fixture.manager.stop()
      expect(stalled.writableEnded).toBe(true)

      await vi.advanceTimersByTimeAsync(20)
      expect(fixture.server.kills).toEqual(['SIGTERM', 'SIGKILL'])
      fixture.server.exited(null, 'SIGKILL')
      await stopping
      await expect(initializing).rejects.toMatchObject({ code: 'app_server_closed' })
      expect(fixture.removeDirectories).toHaveBeenCalledOnce()
    } finally {
      vi.useRealTimers()
    }
  })
})
