import { isAbsolute, resolve } from 'node:path'
import { EnhancedModeComponentManager } from '../packages/codex-bridge/src/component-manager.js'
import manifest from './codex/manifest.json'

function cacheArgument(argv: readonly string[]): string {
  if (argv.length !== 2 || argv[0] !== '--cache' || !argv[1] || !isAbsolute(argv[1])) {
    throw new Error('enhanced_mode_invalid_cache_argument')
  }
  return resolve(argv[1])
}

async function main(): Promise<void> {
  const manager = new EnhancedModeComponentManager({
    cacheRoot: cacheArgument(process.argv.slice(2)),
    manifest,
  })
  const installed = await manager.install()
  process.stdout.write(`${installed.executablePath}\n`)
}

void main().catch((error: unknown) => {
  const code =
    error && typeof error === 'object' && 'code' in error && typeof error.code === 'string'
      ? error.code
      : 'enhanced_mode_install_failed'
  process.stderr.write(`${code}\n`)
  process.exitCode = 1
})
