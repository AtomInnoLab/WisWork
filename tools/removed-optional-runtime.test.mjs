import assert from 'node:assert/strict'
import { existsSync, readFileSync } from 'node:fs'
import { join } from 'node:path'
import test from 'node:test'

const root = new URL('..', import.meta.url).pathname

const removedPaths = [
  'packages/codex-bridge',
  'packages/agent-runtime',
  'tools/codex',
  'tools/install-enhanced-component.ts',
  'tools/run-enhanced-mode-evals.mjs',
  'tools/codex-release.test.mjs',
  'apps/shell/src/main/codex-ipc.ts',
  'apps/shell/src/main/codex-runtime.ts',
  'apps/shell/src/main/enhanced-mode-component.ts',
  'apps/shell/src/shared/codex-api.ts',
  'apps/shell/src/shared/enhanced-mode-api.ts',
  'apps/shell/src/renderer/src/enhanced-mode-view.ts',
  'apps/latex/src/renderer/ai/codex-tool-session.ts',
  'docs/releases/enhanced-mode-pilot.md',
]

const activeFiles = [
  'package.json',
  'apps/shell/package.json',
  'apps/latex/package.json',
  '.github/workflows/package-macos.yml',
  'apps/shell/src/main/index.ts',
  'apps/shell/src/preload/index.ts',
  'apps/shell/src/renderer/src/Home.tsx',
  'apps/latex/src/preload/index.ts',
  'apps/latex/src/renderer/ai/AiPanel.tsx',
  'tools/check-licenses.mjs',
  'tools/gen-third-party-notices.mjs',
]

test('Codex and Enhanced mode are absent from the active product and release surface', () => {
  assert.deepEqual(
    removedPaths.filter((path) => existsSync(join(root, path))),
    [],
  )

  const references = activeFiles.flatMap((path) => {
    const source = readFileSync(join(root, path), 'utf8')
    return /\bcodex\b|enhanced[ _-]?mode|@wiswork\/(?:codex-bridge|agent-runtime)/i.test(source)
      ? [path]
      : []
  })
  assert.deepEqual(references, [])
})
