import { spawn as nodeSpawn, type SpawnOptions } from 'node:child_process'
import { mkdir, mkdtemp, rm } from 'node:fs/promises'
import { basename, isAbsolute, join, relative, resolve, sep } from 'node:path'
import { tmpdir } from 'node:os'
import type { EventEmitter } from 'node:events'
import type { Readable, Writable } from 'node:stream'
import { CodexAppServerClient } from './app-server-client.js'
import { CODEX_CLI_VERSION } from './generated/index.js'
import { JsonRpcClient } from './json-rpc.js'

const VERSION_OUTPUT_LIMIT = 65_536
const STDERR_LIMIT = 65_536
const DEFAULT_STARTUP_TIMEOUT_MS = 5_000
const DEFAULT_GRACEFUL_SHUTDOWN_MS = 500
const DEFAULT_TERMINATE_TIMEOUT_MS = 2_000
const DEFAULT_KILL_TIMEOUT_MS = 2_000
const MAX_BRIDGE_SECRET_BYTES = 4_096
const MAX_DEVELOPER_INSTRUCTIONS_BYTES = 65_536

export interface ChildProcessAdapter extends EventEmitter {
  readonly stdin: Writable
  readonly stdout: Readable
  readonly stderr: Readable
  readonly killed: boolean
  readonly exitCode: number | null
  readonly signalCode: NodeJS.Signals | null
  kill(signal?: NodeJS.Signals): boolean
}

export interface CodexSpawnOptions {
  readonly cwd: string
  readonly env: NodeJS.ProcessEnv
  readonly stdio: 'pipe'
  readonly windowsHide: true
}

export type CodexSpawn = (
  executable: string,
  args: readonly string[],
  options: CodexSpawnOptions,
) => ChildProcessAdapter

export interface OwnedCodexDirectories {
  readonly root: string
  readonly codexHome: string
  readonly cwd: string
}

export interface CodexProcessDiagnostic {
  readonly code: string
}

export interface CodexProcessManagerOptions {
  readonly executablePath: string
  readonly bridge: { readonly baseUrl: string; readonly secret: string }
  readonly mcp?: { readonly url: string; readonly secret: string }
  readonly developerInstructions: string
  readonly spawn?: CodexSpawn
  readonly createDirectories?: () => Promise<OwnedCodexDirectories>
  readonly removeDirectories?: (directories: OwnedCodexDirectories) => Promise<void>
  readonly diagnostics?: (diagnostic: CodexProcessDiagnostic) => void
  readonly startupTimeoutMs?: number
  readonly gracefulShutdownMs?: number
  readonly terminateTimeoutMs?: number
  readonly killTimeoutMs?: number
}

export class CodexProcessError extends Error {
  constructor(readonly code: string) {
    super(code)
    this.name = 'CodexProcessError'
  }
}

function duration(value: number | undefined, fallback: number, code: string): number {
  const resolved = value ?? fallback
  if (!Number.isSafeInteger(resolved) || resolved <= 0) throw new TypeError(code)
  return resolved
}

function validateBridge(baseUrl: string, secret: string): string {
  try {
    const url = new URL(baseUrl)
    if (
      url.protocol !== 'http:' ||
      url.hostname !== '127.0.0.1' ||
      url.port === '' ||
      (url.pathname !== '' && url.pathname !== '/') ||
      url.username !== '' ||
      url.password !== '' ||
      url.search !== '' ||
      url.hash !== '' ||
      typeof secret !== 'string' ||
      secret === '' ||
      Buffer.byteLength(secret, 'utf8') > MAX_BRIDGE_SECRET_BYTES
    ) {
      throw new Error()
    }
    return url.origin
  } catch {
    throw new TypeError('invalid_codex_bridge')
  }
}

function validateMcp(
  mcp: CodexProcessManagerOptions['mcp'],
): { readonly url: string; readonly secret: string } | undefined {
  if (mcp === undefined) return undefined
  try {
    const url = new URL(mcp.url)
    if (
      url.protocol !== 'http:' ||
      url.hostname !== '127.0.0.1' ||
      url.port === '' ||
      !/^\/mcp\/[A-Za-z0-9_-]{43}$/.test(url.pathname) ||
      url.username !== '' ||
      url.password !== '' ||
      url.search !== '' ||
      url.hash !== '' ||
      typeof mcp.secret !== 'string' ||
      !/^[A-Za-z0-9_-]{43}$/.test(mcp.secret) ||
      Buffer.from(mcp.secret, 'base64url').length !== 32 ||
      Buffer.from(mcp.secret, 'base64url').toString('base64url') !== mcp.secret
    ) {
      throw new Error()
    }
    return { url: url.href, secret: mcp.secret }
  } catch {
    throw new TypeError('invalid_codex_mcp')
  }
}

async function defaultCreateDirectories(): Promise<OwnedCodexDirectories> {
  const root = await mkdtemp(join(tmpdir(), 'wiswork-codex-'))
  const codexHome = join(root, 'home')
  const cwd = join(root, 'workspace')
  await mkdir(codexHome, { mode: 0o700 })
  await mkdir(cwd, { mode: 0o700 })
  return { root, codexHome, cwd }
}

async function defaultRemoveDirectories(directories: OwnedCodexDirectories): Promise<void> {
  await rm(directories.root, { recursive: true, force: true })
}

const defaultSpawn: CodexSpawn = (executable, args, options) =>
  nodeSpawn(executable, args, options as SpawnOptions) as unknown as ChildProcessAdapter

function platformEssentials(): NodeJS.ProcessEnv {
  if (process.platform !== 'win32') return {}
  const allowed = ['SystemRoot', 'WINDIR', 'COMSPEC', 'PATHEXT', 'TEMP', 'TMP'] as const
  return Object.fromEntries(
    allowed.flatMap((key) => (process.env[key] === undefined ? [] : [[key, process.env[key]]])),
  )
}

function minimalEnvironment(
  directories: OwnedCodexDirectories,
  token?: string,
  mcpToken?: string,
): NodeJS.ProcessEnv {
  return {
    ...platformEssentials(),
    CODEX_HOME: directories.codexHome,
    ...(token === undefined ? {} : { WISWORK_CODEX_TOKEN: token }),
    ...(mcpToken === undefined ? {} : { WISWORK_MCP_TOKEN: mcpToken }),
    NO_PROXY: '127.0.0.1,localhost',
    no_proxy: '127.0.0.1,localhost',
  }
}

function serverArguments(baseUrl: string, mcpUrl?: string): readonly string[] {
  const configs = [
    'model_provider="wiswork"',
    'model_providers.wiswork.name="WisWork"',
    `model_providers.wiswork.base_url=${JSON.stringify(`${baseUrl}/v1`)}`,
    'model_providers.wiswork.env_key="WISWORK_CODEX_TOKEN"',
    'model_providers.wiswork.wire_api="responses"',
    'features.shell_tool=false',
    'features.unified_exec=false',
    'tools.update_plan.enabled=false',
    'features.multi_agent=false',
    ...(mcpUrl === undefined
      ? []
      : [
          `mcp_servers.wiswork.url=${JSON.stringify(mcpUrl)}`,
          'mcp_servers.wiswork.bearer_token_env_var="WISWORK_MCP_TOKEN"',
        ]),
  ]
  return ['app-server', '--strict-config', '--stdio', ...configs.flatMap((value) => ['-c', value])]
}

interface ExitResult {
  readonly code: number | null
  readonly signal: NodeJS.Signals | null
}

export class CodexProcessManager {
  readonly #executablePath: string
  readonly #baseUrl: string
  readonly #secret: string
  readonly #mcp: { readonly url: string; readonly secret: string } | undefined
  readonly #developerInstructions: string
  readonly #spawn: CodexSpawn
  readonly #createDirectories: () => Promise<OwnedCodexDirectories>
  readonly #removeDirectories: (directories: OwnedCodexDirectories) => Promise<void>
  readonly #diagnostics?: (diagnostic: CodexProcessDiagnostic) => void
  readonly #startupTimeoutMs: number
  readonly #gracefulShutdownMs: number
  readonly #terminateTimeoutMs: number
  readonly #killTimeoutMs: number
  readonly crashed: Promise<CodexProcessError>
  readonly #resolveCrashed: (error: CodexProcessError) => void
  #state: 'new' | 'starting' | 'running' | 'stopping' | 'stopped' | 'failed' = 'new'
  #startPromise: Promise<CodexAppServerClient> | undefined
  #stopPromise: Promise<void> | undefined
  #directories: OwnedCodexDirectories | undefined
  #cleanupPromise: Promise<void> | undefined
  #child: ChildProcessAdapter | undefined
  #client: CodexAppServerClient | undefined
  #exitPromise: Promise<ExitResult> | undefined
  #resolveExit: ((result: ExitResult) => void) | undefined
  #unexpectedExit: CodexProcessError | undefined
  #canCleanup = true

  constructor(options: CodexProcessManagerOptions) {
    if (typeof options?.executablePath !== 'string' || !isAbsolute(options.executablePath)) {
      throw new TypeError('codex_executable_must_be_absolute')
    }
    if (
      typeof options.developerInstructions !== 'string' ||
      options.developerInstructions === '' ||
      Buffer.byteLength(options.developerInstructions, 'utf8') > MAX_DEVELOPER_INSTRUCTIONS_BYTES
    ) {
      throw new TypeError('invalid_developer_instructions')
    }
    if ((options.createDirectories === undefined) !== (options.removeDirectories === undefined)) {
      throw new TypeError('codex_temp_adapters_must_be_paired')
    }
    this.#executablePath = options.executablePath
    this.#baseUrl = validateBridge(options.bridge?.baseUrl, options.bridge?.secret)
    this.#secret = options.bridge.secret
    this.#mcp = validateMcp(options.mcp)
    this.#developerInstructions = options.developerInstructions
    this.#spawn = options.spawn ?? defaultSpawn
    this.#createDirectories = options.createDirectories ?? defaultCreateDirectories
    this.#removeDirectories = options.removeDirectories ?? defaultRemoveDirectories
    this.#diagnostics = options.diagnostics
    this.#startupTimeoutMs = duration(
      options.startupTimeoutMs,
      DEFAULT_STARTUP_TIMEOUT_MS,
      'invalid_startup_timeout',
    )
    this.#gracefulShutdownMs = duration(
      options.gracefulShutdownMs,
      DEFAULT_GRACEFUL_SHUTDOWN_MS,
      'invalid_shutdown_timeout',
    )
    this.#terminateTimeoutMs = duration(
      options.terminateTimeoutMs,
      DEFAULT_TERMINATE_TIMEOUT_MS,
      'invalid_terminate_timeout',
    )
    this.#killTimeoutMs = duration(
      options.killTimeoutMs,
      DEFAULT_KILL_TIMEOUT_MS,
      'invalid_kill_timeout',
    )
    let resolveCrashed!: (error: CodexProcessError) => void
    this.crashed = new Promise((resolve) => {
      resolveCrashed = resolve
    })
    this.#resolveCrashed = resolveCrashed
  }

  start(): Promise<CodexAppServerClient> {
    if (this.#state === 'starting') return this.#startPromise!
    if (this.#state !== 'new') {
      return Promise.reject(new CodexProcessError('codex_process_already_started'))
    }
    this.#state = 'starting'
    this.#startPromise = this.#startInternal()
    return this.#startPromise
  }

  stop(): Promise<void> {
    if (this.#stopPromise) return this.#stopPromise
    this.#stopPromise = this.#stopInternal()
    return this.#stopPromise
  }

  async #startInternal(): Promise<CodexAppServerClient> {
    try {
      const directories = await this.#createDirectories()
      this.#validateDirectories(directories)
      this.#directories = directories
      await this.#verifyVersion(this.#directories)
      if (this.#state !== 'starting') throw new CodexProcessError('codex_process_start_failed')
      const child = this.#spawn(
        this.#executablePath,
        serverArguments(this.#baseUrl, this.#mcp?.url),
        {
          cwd: this.#directories.cwd,
          env: minimalEnvironment(this.#directories, this.#secret, this.#mcp?.secret),
          stdio: 'pipe',
          windowsHide: true,
        },
      )
      this.#child = child
      this.#installStderrDrain(child.stderr)
      this.#exitPromise = new Promise((resolve) => {
        this.#resolveExit = resolve
      })
      child.once('exit', this.#onServerExit)
      child.once('error', this.#onServerRuntimeError)
      await this.#waitForSpawn(child)
      if (this.#unexpectedExit) throw this.#unexpectedExit
      const rpc = new JsonRpcClient({ input: child.stdout, output: child.stdin })
      this.#client = new CodexAppServerClient({
        rpc,
        cwd: this.#directories.cwd,
        developerInstructions: this.#developerInstructions,
      })
      this.#state = 'running'
      return this.#client
    } catch (error) {
      let failure =
        error instanceof CodexProcessError
          ? error
          : new CodexProcessError('codex_process_start_failed')
      const child = this.#child
      if (child && child.exitCode === null && child.signalCode === null) {
        this.#state = 'stopping'
        child.kill('SIGKILL')
        if (!(await this.#waitForExit(this.#killTimeoutMs))) {
          this.#canCleanup = false
          failure = new CodexProcessError('codex_process_termination_timeout')
        }
      }
      this.#state = 'failed'
      if (this.#canCleanup) await this.#cleanup()
      throw failure
    }
  }

  #validateDirectories(directories: OwnedCodexDirectories): void {
    const normalizedTemporaryRoot = resolve(tmpdir())
    const normalizedRoot = resolve(directories.root)
    const temporaryRelative = relative(normalizedTemporaryRoot, normalizedRoot)
    const isOwnedChild = (path: string): boolean => {
      const child = relative(directories.root, path)
      return child !== '' && child !== '..' && !child.startsWith(`..${sep}`) && !isAbsolute(child)
    }
    if (
      !isAbsolute(directories.root) ||
      !isAbsolute(directories.codexHome) ||
      !isAbsolute(directories.cwd) ||
      normalizedRoot !== directories.root ||
      temporaryRelative === '' ||
      temporaryRelative === '..' ||
      temporaryRelative.startsWith(`..${sep}`) ||
      isAbsolute(temporaryRelative) ||
      !basename(normalizedRoot).startsWith('wiswork-codex-') ||
      directories.codexHome !== join(normalizedRoot, 'home') ||
      directories.cwd !== join(normalizedRoot, 'workspace') ||
      !isOwnedChild(directories.codexHome) ||
      !isOwnedChild(directories.cwd)
    ) {
      throw new CodexProcessError('invalid_owned_directories')
    }
  }

  #verifyVersion(directories: OwnedCodexDirectories): Promise<void> {
    return new Promise((resolve, reject) => {
      let child: ChildProcessAdapter
      try {
        child = this.#spawn(this.#executablePath, ['--version'], {
          cwd: directories.cwd,
          env: minimalEnvironment(directories),
          stdio: 'pipe',
          windowsHide: true,
        })
      } catch {
        reject(new CodexProcessError('codex_version_check_failed'))
        return
      }
      let stdout = Buffer.alloc(0)
      let outputBytes = 0
      let settled = false
      let terminationError: CodexProcessError | undefined
      let terminationTimer: ReturnType<typeof setTimeout> | undefined
      const cleanup = (): void => {
        clearTimeout(timer)
        if (terminationTimer !== undefined) clearTimeout(terminationTimer)
        child.stdout.off('data', onStdout)
        child.stderr.off('data', onStderr)
        child.off('error', onError)
        child.off('exit', onExit)
      }
      const finish = (error?: CodexProcessError): void => {
        if (settled) return
        settled = true
        cleanup()
        if (error) reject(error)
        else resolve()
      }
      const count = (chunk: Buffer | string): boolean => {
        outputBytes += Buffer.byteLength(chunk)
        if (outputBytes <= VERSION_OUTPUT_LIMIT) return true
        terminate(new CodexProcessError('codex_version_output_limit'))
        return false
      }
      const onStdout = (chunk: Buffer | string): void => {
        if (!count(chunk)) return
        stdout = Buffer.concat([stdout, Buffer.isBuffer(chunk) ? chunk : Buffer.from(chunk)])
      }
      const onStderr = (chunk: Buffer | string): void => {
        count(chunk)
      }
      const onError = (): void => {
        if (terminationError) return
        finish(new CodexProcessError('codex_version_check_failed'))
      }
      const onExit = (code: number | null): void => {
        if (settled) return
        if (terminationError) {
          finish(terminationError)
          return
        }
        if (code !== 0) {
          finish(new CodexProcessError('codex_version_check_failed'))
          return
        }
        let version: string
        try {
          version = new TextDecoder('utf-8', { fatal: true }).decode(stdout).trim()
        } catch {
          finish(new CodexProcessError('codex_version_mismatch'))
          return
        }
        finish(
          version === CODEX_CLI_VERSION
            ? undefined
            : new CodexProcessError('codex_version_mismatch'),
        )
      }
      const terminate = (error: CodexProcessError): void => {
        if (settled || terminationError) return
        terminationError = error
        clearTimeout(timer)
        child.stdout.off('data', onStdout)
        child.stderr.off('data', onStderr)
        child.kill('SIGKILL')
        terminationTimer = setTimeout(() => {
          this.#canCleanup = false
          finish(new CodexProcessError('codex_process_termination_timeout'))
        }, this.#killTimeoutMs)
        terminationTimer.unref()
      }
      const timer = setTimeout(() => {
        terminate(new CodexProcessError('codex_version_timeout'))
      }, this.#startupTimeoutMs)
      timer.unref()
      child.stdout.on('data', onStdout)
      child.stderr.on('data', onStderr)
      child.once('error', onError)
      child.once('exit', onExit)
    })
  }

  #waitForSpawn(child: ChildProcessAdapter): Promise<void> {
    return new Promise((resolve, reject) => {
      let settled = false
      const cleanup = (): void => {
        clearTimeout(timer)
        child.off('spawn', onSpawn)
        child.off('error', onError)
        child.off('exit', onExit)
      }
      const finish = (error?: CodexProcessError): void => {
        if (settled) return
        settled = true
        cleanup()
        if (error) reject(error)
        else resolve()
      }
      const onSpawn = (): void => finish()
      const onError = (): void => finish(new CodexProcessError('codex_process_start_failed'))
      const onExit = (): void => finish(new CodexProcessError('codex_process_exited'))
      const timer = setTimeout(
        () => finish(new CodexProcessError('codex_process_start_timeout')),
        this.#startupTimeoutMs,
      )
      timer.unref()
      child.once('spawn', onSpawn)
      child.once('error', onError)
      child.once('exit', onExit)
    })
  }

  #installStderrDrain(stderr: Readable): void {
    let bytes = 0
    let reported = false
    let truncated = false
    stderr.on('data', (chunk: Buffer | string) => {
      if (!reported) {
        reported = true
        this.#emitDiagnostic('codex_stderr_output')
      }
      bytes += Buffer.byteLength(chunk)
      if (!truncated && bytes > STDERR_LIMIT) {
        truncated = true
        this.#emitDiagnostic('codex_stderr_truncated')
      }
    })
    stderr.resume()
  }

  readonly #onServerRuntimeError = (): void => {
    if (this.#state === 'stopping' || this.#state === 'stopped') return
    this.#handleUnexpectedExit(new CodexProcessError('codex_process_exited'))
  }

  readonly #onServerExit = (code: number | null, signal: NodeJS.Signals | null): void => {
    this.#resolveExit?.({ code, signal })
    this.#resolveExit = undefined
    if (this.#state === 'stopping' || this.#state === 'stopped') return
    this.#handleUnexpectedExit(new CodexProcessError('codex_process_exited'))
  }

  #handleUnexpectedExit(error: CodexProcessError): void {
    if (this.#unexpectedExit) return
    this.#unexpectedExit = error
    this.#state = 'failed'
    void (async () => {
      void this.#client?.shutdown(error.code)
      try {
        await this.#cleanup()
      } catch {
        // Cleanup reports its own redacted diagnostic; crash reporting must still settle.
      } finally {
        this.#emitDiagnostic(error.code)
        this.#resolveCrashed(error)
      }
    })()
  }

  async #stopInternal(): Promise<void> {
    if (this.#state === 'new') {
      this.#state = 'stopped'
      return
    }
    if (this.#state === 'starting') {
      try {
        await this.#startPromise
      } catch {
        await this.#cleanup()
        return
      }
    }
    if (this.#state === 'failed' || this.#state === 'stopped') {
      await this.#cleanup()
      return
    }
    this.#state = 'stopping'
    void this.#client?.shutdown()
    const child = this.#child
    if (child && !(await this.#waitForExit(this.#gracefulShutdownMs))) {
      child.kill('SIGTERM')
      if (!(await this.#waitForExit(this.#terminateTimeoutMs))) {
        child.kill('SIGKILL')
        if (!(await this.#waitForExit(this.#killTimeoutMs))) {
          this.#state = 'failed'
          throw new CodexProcessError('codex_process_stop_timeout')
        }
      }
    }
    this.#state = 'stopped'
    await this.#cleanup()
  }

  async #waitForExit(timeoutMs: number): Promise<boolean> {
    const child = this.#child
    if (!child || child.exitCode !== null || child.signalCode !== null) return true
    const exited = this.#exitPromise
    if (!exited) return true
    return new Promise((resolve) => {
      const timer = setTimeout(() => resolve(false), timeoutMs)
      timer.unref()
      void exited.then(() => {
        clearTimeout(timer)
        resolve(true)
      })
    })
  }

  #cleanup(): Promise<void> {
    if (this.#cleanupPromise) return this.#cleanupPromise
    if (!this.#canCleanup) return Promise.reject(new CodexProcessError('codex_process_alive'))
    if (!this.#directories) return Promise.resolve()
    const directories = this.#directories
    this.#cleanupPromise = this.#removeDirectories(directories).catch(() => {
      this.#emitDiagnostic('codex_temp_cleanup_failed')
      throw new CodexProcessError('codex_temp_cleanup_failed')
    })
    return this.#cleanupPromise
  }

  #emitDiagnostic(code: string): void {
    try {
      this.#diagnostics?.({ code })
    } catch {
      // Diagnostic sinks cannot affect process lifecycle.
    }
  }
}
