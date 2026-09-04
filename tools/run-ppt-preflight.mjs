import { createHash } from 'node:crypto'
import { existsSync, mkdirSync, readFileSync, rmSync, statSync, writeFileSync } from 'node:fs'
import { homedir } from 'node:os'
import { isAbsolute, join, resolve } from 'node:path'
import { spawn } from 'node:child_process'
import { fileURLToPath } from 'node:url'

const npm = process.platform === 'win32' ? 'npm.cmd' : 'npm'
const npx = process.platform === 'win32' ? 'npx.cmd' : 'npx'
const VERSION = '0.147.0'

function platformAsset(platform, arch) {
  if (platform === 'darwin' && arch === 'arm64') return ['darwin-arm64', 'bin/codex-app-server']
  if (platform === 'darwin' && arch === 'x64') return ['darwin-x64', 'bin/codex-app-server']
  if (platform === 'win32' && arch === 'x64') return ['win32-x64', 'bin/codex-app-server.exe']
  return undefined
}

export function resolvePreflightExecutable({ env, platform, arch, home = homedir() }) {
  if (env.WISWORK_CODEX_INTEGRATION_EXECUTABLE) {
    return resolve(env.WISWORK_CODEX_INTEGRATION_EXECUTABLE)
  }
  const asset = platformAsset(platform, arch)
  const candidates = []
  if (asset) {
    const userData =
      platform === 'darwin'
        ? join(home, 'Library', 'Application Support', 'WisWork')
        : join(env.APPDATA ?? join(home, 'AppData', 'Roaming'), 'WisWork')
    candidates.push(join(userData, 'components', 'enhanced-mode', VERSION, asset[0], asset[1]))
  }
  if (platform === 'linux' && arch === 'x64') {
    candidates.push(
      join(
        home,
        '.codex',
        'packages',
        'standalone',
        'releases',
        '0.147.0-x86_64-unknown-linux-musl',
        'bin',
        'codex',
      ),
    )
  }
  return candidates.find((candidate) => existsSync(candidate))
}

export function createPptPreflightSteps({ env, platform, executable, artifactRoot }) {
  if (!env.WISWORK_REAL_WISUSAGE_TOKEN) throw new Error('ppt_preflight_wisusage_token_required')
  if (
    !executable ||
    !isAbsolute(executable) ||
    !existsSync(executable) ||
    !statSync(executable).isFile()
  )
    throw new Error('ppt_preflight_codex_component_required')
  const standardOutput = join(artifactRoot, 'standard-llm.pptx')
  const enhancedOutput = join(artifactRoot, 'enhanced-onboarding.pptx')
  const sharedEnv = {
    WISWORK_REAL_WISUSAGE_TOKEN: env.WISWORK_REAL_WISUSAGE_TOKEN,
    WISWORK_CODEX_INTEGRATION_EXECUTABLE: executable,
  }
  const electron = [npm, ['run', 'test:e2e', '--', 'e2e/slides-acceptance-render.spec.ts']]
  return {
    artifacts: { standardOutput, enhancedOutput },
    steps: [
      [
        'production-contracts',
        npx,
        [
          'vitest',
          'run',
          'packages/codex-bridge/tests/task4-review-hardening.test.ts',
          'apps/shell/tests/pc-codex-hosts.test.ts',
          'packages/agent-runtime/tests/renderer-snapshot.test.ts',
          'apps/slides/tests/build-deck-tool.test.ts',
        ],
        {},
      ],
      [
        'standard-live-ppt',
        npx,
        ['vitest', 'run', 'apps/slides/tests/standard-live-ppt.integration.test.ts'],
        { ...sharedEnv, WISWORK_STANDARD_PPT_E2E_OUTPUT: standardOutput },
      ],
      [
        'enhanced-live-ppt',
        npx,
        [
          'vitest',
          'run',
          'apps/shell/tests/codex-engine.integration.test.ts',
          '-t',
          'accepts the live WisUsage stream shape|generates and verifies a three-page onboarding deck',
        ],
        { ...sharedEnv, WISWORK_ENHANCED_PPT_E2E_OUTPUT: enhancedOutput },
      ],
      [
        'electron-source-e2e',
        platform === 'linux' && !env.DISPLAY ? 'xvfb-run' : electron[0],
        platform === 'linux' && !env.DISPLAY
          ? ['--auto-servernum', '--', electron[0], ...electron[1]]
          : electron[1],
        {},
      ],
    ],
  }
}

function inspectPptx(path) {
  const bytes = readFileSync(path)
  if (bytes.length < 1_024 || bytes[0] !== 0x50 || bytes[1] !== 0x4b)
    throw new Error('ppt_preflight_artifact_invalid')
  return {
    path,
    bytes: bytes.length,
    sha256: createHash('sha256').update(bytes).digest('hex'),
  }
}

export async function orchestratePptPreflight({ runStep, env, platform, arch, home, root }) {
  const artifactRoot = resolve(root, 'test-results', 'ppt-preflight')
  rmSync(artifactRoot, { recursive: true, force: true })
  mkdirSync(artifactRoot, { recursive: true })
  const executable = resolvePreflightExecutable({ env, platform, arch, home })
  const definition = createPptPreflightSteps({ env, platform, executable, artifactRoot })
  const completed = []
  const startedAt = Date.now()
  for (const [name, command, args, extraEnv] of definition.steps) {
    const stepStartedAt = Date.now()
    const code = await runStep({ name, command, args, extraEnv })
    completed.push({ name, code, durationMs: Date.now() - stepStartedAt })
    if (code !== 0) {
      writeFileSync(
        join(artifactRoot, 'report.json'),
        `${JSON.stringify(
          {
            schema: 'wiswork-ppt-preflight/v1',
            status: 'failed',
            failedStep: name,
            generatedAt: Date.now(),
            durationMs: Date.now() - startedAt,
            componentVersion: VERSION,
            steps: completed,
          },
          null,
          2,
        )}\n`,
        { mode: 0o600 },
      )
      return code
    }
  }
  const report = {
    schema: 'wiswork-ppt-preflight/v1',
    status: 'passed',
    generatedAt: Date.now(),
    durationMs: Date.now() - startedAt,
    componentVersion: VERSION,
    steps: completed,
    artifacts: [
      inspectPptx(definition.artifacts.standardOutput),
      inspectPptx(definition.artifacts.enhancedOutput),
    ],
  }
  writeFileSync(join(artifactRoot, 'report.json'), `${JSON.stringify(report, null, 2)}\n`, {
    mode: 0o600,
  })
  return 0
}

function spawnStep({ command, args, extraEnv }) {
  return new Promise((resolveCode) => {
    const child = spawn(command, args, {
      cwd: process.cwd(),
      env: { ...process.env, ...extraEnv },
      stdio: 'inherit',
      shell: false,
    })
    child.once('error', () => resolveCode(1))
    child.once('exit', (code) => resolveCode(code ?? 1))
  })
}

if (process.argv[1] && fileURLToPath(import.meta.url) === resolve(process.argv[1])) {
  try {
    process.exitCode = await orchestratePptPreflight({
      runStep: spawnStep,
      env: process.env,
      platform: process.platform,
      arch: process.arch,
      home: homedir(),
      root: process.cwd(),
    })
  } catch (error) {
    const code = error instanceof Error ? error.message : 'ppt_preflight_failed'
    console.error(code)
    process.exitCode = 1
  }
}
