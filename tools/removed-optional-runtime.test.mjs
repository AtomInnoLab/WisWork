import assert from 'node:assert/strict'
import { existsSync, mkdtempSync, readFileSync, readdirSync, rmSync, writeFileSync } from 'node:fs'
import { tmpdir } from 'node:os'
import { extname, join } from 'node:path'
import test from 'node:test'
import { fileURLToPath } from 'node:url'

const root = fileURLToPath(new URL('..', import.meta.url))
const forbiddenRuntime = /\bcodex\b|enhanced[ _-]?mode|@wiswork\/(?:codex-bridge|agent-runtime)/i
const sourceExtensions = new Set([
  '.cjs',
  '.css',
  '.html',
  '.js',
  '.json',
  '.mjs',
  '.ts',
  '.tsx',
  '.yaml',
  '.yml',
])
const excludedDirectories = new Set(['build', 'dist', 'node_modules', 'out', 'release', 'target'])
const excludedFiles = new Set(['tools/removed-optional-runtime.test.mjs'])

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

const activeRoots = [
  'apps',
  'packages',
  'tools',
  '.github/workflows',
  'package.json',
  'package-lock.json',
]

function activeSourceFiles(base, relative = '') {
  const absolute = join(base, relative)
  if (!existsSync(absolute)) return []

  const entries = readdirSync(absolute, { withFileTypes: true })
  return entries.flatMap((entry) => {
    const child = join(relative, entry.name)
    if (entry.isDirectory()) {
      return excludedDirectories.has(entry.name) ? [] : activeSourceFiles(base, child)
    }
    return entry.isFile() && sourceExtensions.has(extname(entry.name)) ? [child] : []
  })
}

function forbiddenReferences(base, roots) {
  return roots.flatMap((path) => {
    const absolute = join(base, path)
    const files = existsSync(absolute) && !extname(path) ? activeSourceFiles(base, path) : [path]
    return files.flatMap((file) => {
      if (excludedFiles.has(file)) return []
      return forbiddenRuntime.test(readFileSync(join(base, file), 'utf8')) ? [file] : []
    })
  })
}

test('Codex and Enhanced mode are absent from the active product and release surface', () => {
  assert.deepEqual(
    removedPaths.filter((path) => existsSync(join(root, path))),
    [],
  )

  assert.deepEqual(forbiddenReferences(root, activeRoots), [])
})

test('the removal policy detects a forbidden runtime in a newly added source file', () => {
  const fixture = mkdtempSync(join(tmpdir(), 'wiswork-runtime-policy-'))
  try {
    writeFileSync(join(fixture, 'new-runtime.ts'), "export const runtime = 'codex'\n")
    assert.deepEqual(forbiddenReferences(fixture, ['new-runtime.ts']), ['new-runtime.ts'])
  } finally {
    rmSync(fixture, { recursive: true, force: true })
  }
})
