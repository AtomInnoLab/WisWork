import { spawn as nodeSpawn, spawnSync, type ChildProcess } from 'node:child_process'
import { access, mkdir, rm, writeFile } from 'node:fs/promises'
import { join, resolve } from 'node:path'
import {
  CodexProcessManager,
  EnhancedModeComponentManager,
  parseEnhancedCapabilityAuthorization,
  type CodexSpawn,
} from '../packages/codex-bridge/src/index.js'
import {
  parseEnhancedRolloutPolicy,
  shouldStartEnhancedRuntime,
} from '../packages/agent-runtime/src/index.js'
import { ShellCodexRuntime } from '../apps/shell/src/main/codex-runtime.js'
import manifest from './codex/manifest.json'

async function main(): Promise<void> {
  const index = process.argv.indexOf('--cache')
  if (index < 0 || !process.argv[index + 1]) throw new Error('native_lifecycle_cache_required')
  const cacheRoot = resolve(process.argv[index + 1])
  const manager = new EnhancedModeComponentManager({ cacheRoot, manifest })

  if ((await manager.status()).state !== 'missing') throw new Error('native_lifecycle_not_clean')

  // Model a genuine upgrade from a previously installed component. Production installation owns
  // version directories and must remove this historical fixture only after 0.147 is promoted.
  const historicalVersion = join(cacheRoot, '0.146.0')
  await mkdir(historicalVersion, { recursive: true })
  await writeFile(join(historicalVersion, 'installed-version'), '0.146.0', { mode: 0o600 })
  const installed = await manager.install()
  if ((await manager.status()).state !== 'ready') throw new Error('native_lifecycle_install_failed')
  try {
    await access(historicalVersion)
    throw new Error('native_lifecycle_update_did_not_retire_old_version')
  } catch (error) {
    if (
      error instanceof Error &&
      error.message === 'native_lifecycle_update_did_not_retire_old_version'
    )
      throw error
  }
  const executable = await manager.resolveExecutable()
  if (executable !== installed.executablePath) throw new Error('native_lifecycle_identity_drift')
  const version = spawnSync(executable, ['--version'], { encoding: 'utf8', timeout: 10_000 })
  if (version.status !== 0 || !/^codex-app-server 0\.147\.0\s*$/.test(version.stdout))
    throw new Error('native_lifecycle_launch_failed')

  // Installed/offline launch verification must not touch either download source.
  const offline = new EnhancedModeComponentManager({
    cacheRoot,
    manifest,
    fetchImplementation: async () => {
      throw new Error('network_forbidden')
    },
  })
  if ((await offline.resolveExecutable()) !== executable)
    throw new Error('native_lifecycle_offline_failed')
  // Installing the already pinned version is the update/no-op path and must retain identity.
  if ((await offline.install()).executablePath !== executable)
    throw new Error('native_lifecycle_noop_failed')

  // Exercise the production process manager with the downloaded native binary. A graceful restart
  // must allocate a distinct OS process and private session directory; a crash must clean that
  // directory before a subsequent relaunch.
  const serverChildren: ChildProcess[] = []
  const sessionRoots: string[] = []
  const spawn: CodexSpawn = (command, args, options) => {
    const child = nodeSpawn(command, args, options)
    // The verified optional component is the dedicated codex-app-server binary, so unlike the
    // full codex CLI it does not receive an `app-server` argv prefix. This injected spawn is used
    // only by CodexProcessManager; every child observed here is therefore the server instance.
    serverChildren.push(child)
    return child
  }
  const createProcess = () =>
    new CodexProcessManager({
      executablePath: executable,
      bridge: { baseUrl: 'http://127.0.0.1:9', secret: 'native-lifecycle-secret' },
      developerInstructions: 'Native release lifecycle verification.',
      spawn,
      createDirectories: async () => {
        const { mkdtemp } = await import('node:fs/promises')
        const { tmpdir } = await import('node:os')
        const root = await mkdtemp(join(tmpdir(), 'wiswork-codex-native-'))
        const codexHome = join(root, 'home')
        const cwd = join(root, 'workspace')
        await mkdir(codexHome, { mode: 0o700 })
        await mkdir(cwd, { mode: 0o700 })
        sessionRoots.push(root)
        return { root, codexHome, cwd }
      },
      removeDirectories: async (directories) => {
        await rm(directories.root, { recursive: true, force: true })
      },
    })

  const first = createProcess()
  await first.start()
  const firstPid = serverChildren.at(-1)?.pid
  await first.stop()
  const restarted = createProcess()
  await restarted.start()
  const restartedPid = serverChildren.at(-1)?.pid
  if (
    !firstPid ||
    !restartedPid ||
    firstPid === restartedPid ||
    sessionRoots[0] === sessionRoots[1]
  )
    throw new Error('native_lifecycle_restart_identity_reused')
  await restarted.stop()

  const crashing = createProcess()
  await crashing.start()
  const crashedRoot = sessionRoots.at(-1)!
  serverChildren.at(-1)?.kill('SIGKILL')
  await crashing.crashed
  try {
    await access(crashedRoot)
    throw new Error('native_lifecycle_crash_cleanup_failed')
  } catch (error) {
    if (error instanceof Error && error.message === 'native_lifecycle_crash_cleanup_failed')
      throw error
  }
  const relaunched = createProcess()
  await relaunched.start()
  await relaunched.stop()

  // Evaluate every rollout control through the production parser/decision function. These are
  // independent controls: Standard mode, global, per-host and raw Office can each deny activation.
  const hosts = Object.fromEntries(
    ['latex', 'slides', 'docs', 'sheets', 'office-word', 'office-excel', 'office-powerpoint'].map(
      (host) => [host, true],
    ),
  )
  const enabled = parseEnhancedRolloutPolicy({ globalEnabled: true, rawOfficeEnabled: true, hosts })
  if (
    !shouldStartEnhancedRuntime(
      { requested: 'enhanced', active: 'enhanced' },
      enabled,
      'office-word',
    )
  )
    throw new Error('native_lifecycle_enabled_policy_denied')
  for (const denied of [
    [{ requested: 'enhanced', active: 'standard' }, enabled, 'latex'],
    [
      { requested: 'enhanced', active: 'enhanced' },
      parseEnhancedRolloutPolicy({ ...enabled, globalEnabled: false }),
      'latex',
    ],
    [
      { requested: 'enhanced', active: 'enhanced' },
      parseEnhancedRolloutPolicy({ ...enabled, hosts: { ...hosts, slides: false } }),
      'slides',
    ],
  ] as const) {
    if (shouldStartEnhancedRuntime(denied[0], denied[1], denied[2]))
      throw new Error('native_lifecycle_kill_switch_failed')
  }
  parseEnhancedCapabilityAuthorization({
    host: 'office-word',
    policy: enabled,
    declaration: { capabilities: ['raw-office-proposal'] },
  })
  let rawDenied = false
  try {
    parseEnhancedCapabilityAuthorization({
      host: 'office-word',
      policy: parseEnhancedRolloutPolicy({ ...enabled, rawOfficeEnabled: false }),
      declaration: { capabilities: ['raw-office-proposal'] },
    })
  } catch {
    rawDenied = true
  }
  if (!rawDenied) throw new Error('native_lifecycle_raw_office_switch_failed')

  // Prove the same switches at the production Shell startup/status boundary, not only in parsers.
  let bootstrapStarts = 0
  const bootstrap = {
    start: async () => {
      bootstrapStarts += 1
      return {
        startTurn: async () => undefined,
        cancelTurn: async () => undefined,
        closeDocument: async () => undefined,
        close: async () => undefined,
      }
    },
  }
  const shellRuntime = (activeAgentRuntime: 'standard' | 'enhanced', policy: typeof enabled) =>
    new ShellCodexRuntime({
      activeAgentRuntime,
      policy,
      isSignedIn: async () => true,
      resolveExecutable: async () => executable,
      bootstrap,
    })
  const standardRuntime = shellRuntime('standard', enabled)
  await standardRuntime.initialize()
  if (standardRuntime.state !== 'standard' || bootstrapStarts !== 0)
    throw new Error('native_lifecycle_standard_not_inert')
  const globallyDisabled = shellRuntime(
    'enhanced',
    parseEnhancedRolloutPolicy({ ...enabled, globalEnabled: false }),
  )
  await globallyDisabled.initialize().then(
    () => {
      throw new Error('native_lifecycle_global_switch_failed')
    },
    () => undefined,
  )
  if (globallyDisabled.state !== 'failed_safe' || bootstrapStarts !== 0)
    throw new Error('native_lifecycle_global_status_failed')
  const hostDisabled = shellRuntime(
    'enhanced',
    parseEnhancedRolloutPolicy({ ...enabled, hosts: { ...hosts, slides: false } }),
  )
  await hostDisabled.initialize()
  let hostUnavailable = false
  try {
    hostDisabled.registerDocument({
      owner: { isDestroyed: () => false },
      documentId: 'native-slides-document',
      host: 'slides',
      generation: 1,
    })
  } catch {
    hostUnavailable = true
  }
  if (!hostUnavailable) throw new Error('native_lifecycle_host_status_failed')
  const rawDisabled = shellRuntime(
    'enhanced',
    parseEnhancedRolloutPolicy({ ...enabled, rawOfficeEnabled: false }),
  )
  await rawDisabled.initialize()
  if (rawDisabled.createOfficeSessionStatement('office-word')?.raw_office !== false)
    throw new Error('native_lifecycle_raw_tool_status_failed')
  await Promise.all([standardRuntime.close(), hostDisabled.close(), rawDisabled.close()])
  await manager.remove()
  if ((await manager.status()).state !== 'missing')
    throw new Error('native_lifecycle_remove_failed')

  // A corrupt primary must use the pinned official fallback, whose signature is verified natively.
  const fallbackRoot = `${cacheRoot}-fallback`
  const fallback = new EnhancedModeComponentManager({
    cacheRoot: fallbackRoot,
    manifest,
    fetchImplementation: (url, init) =>
      String(url).startsWith('https://downloads.wiswork.com/')
        ? Promise.resolve(new Response('corrupt-primary'))
        : fetch(url, init),
  })
  await fallback.install()
  if ((await fallback.status()).state !== 'ready')
    throw new Error('native_lifecycle_fallback_failed')
  await fallback.remove()

  // Neither source may be accepted when both return corrupt bytes.
  const corrupt = new EnhancedModeComponentManager({
    cacheRoot: `${cacheRoot}-both-corrupt`,
    manifest,
    fetchImplementation: async () => new Response('corrupt-source'),
  })
  let rejected = false
  try {
    await corrupt.install()
  } catch {
    rejected = true
  }
  if (!rejected || (await corrupt.status()).state !== 'missing')
    throw new Error('native_lifecycle_corrupt_fail_open')
}

void main().catch((error: unknown) => {
  const message = error instanceof Error ? error.message : 'native_lifecycle_failed'
  process.stderr.write(`${message}\n`)
  process.exitCode = 1
})
