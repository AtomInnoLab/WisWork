import { spawn } from 'node:child_process'
import { fileURLToPath } from 'node:url'
import { resolve as resolvePath } from 'node:path'

const npm = process.platform === 'win32' ? 'npm.cmd' : 'npm'
const npx = process.platform === 'win32' ? 'npx.cmd' : 'npx'

export async function orchestrateSlidesAcceptanceE2E(
  runStep,
  report = console.error,
  playwrightFilters = [],
) {
  let primaryCode = 0
  let cleanupCode = 0
  const primary = [
    [
      'e2e-build',
      npm,
      ['run', 'build', '-w', '@wiswork/slides'],
      { WISWORK_SLIDES_ACCEPTANCE_E2E: '1' },
    ],
    ['e2e-artifact', process.execPath, ['tools/check-slides-e2e-artifact.mjs', 'present']],
    [
      'playwright',
      npx,
      ['playwright', 'test', '--config', 'e2e/playwright.config.ts', ...playwrightFilters],
    ],
  ]
  const cleanup = [
    ['default-build', npm, ['run', 'build', '-w', '@wiswork/slides']],
    ['default-artifact', process.execPath, ['tools/check-slides-e2e-artifact.mjs', 'absent']],
  ]
  try {
    for (const [name, command, args, extraEnv] of primary) {
      const code = await runStep({ name, command, args, extraEnv })
      if (code !== 0) {
        primaryCode = code
        break
      }
    }
  } finally {
    for (const [name, command, args, extraEnv] of cleanup) {
      const code = await runStep({ name, command, args, extraEnv })
      if (code !== 0 && cleanupCode === 0) cleanupCode = code
    }
    if (cleanupCode !== 0) report(`Slides acceptance E2E cleanup failed (${cleanupCode}).`)
  }
  return primaryCode || cleanupCode
}

function spawnStep({ command, args, extraEnv }) {
  return new Promise((resolve) => {
    const child = spawn(command, args, {
      cwd: process.cwd(),
      env: { ...process.env, ...extraEnv },
      stdio: 'inherit',
      shell: false,
    })
    child.once('error', (error) => {
      console.error(error)
      resolve(1)
    })
    child.once('exit', (code, signal) => {
      if (signal) console.error(`Command terminated by ${signal}: ${command}`)
      resolve(code ?? 1)
    })
  })
}

if (process.argv[1] && fileURLToPath(import.meta.url) === resolvePath(process.argv[1])) {
  process.exitCode = await orchestrateSlidesAcceptanceE2E(
    spawnStep,
    console.error,
    process.argv.slice(2),
  )
}
