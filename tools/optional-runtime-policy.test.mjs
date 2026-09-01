import assert from 'node:assert/strict'
import { existsSync, readFileSync, readdirSync } from 'node:fs'
import { basename, extname, join } from 'node:path'
import test from 'node:test'
import { fileURLToPath } from 'node:url'

const root = fileURLToPath(new URL('..', import.meta.url))
const runtimePackages = ['packages/agent-runtime', 'packages/codex-bridge']
const binaryExtensions = new Set(['', '.app', '.dll', '.dylib', '.exe', '.node', '.so', '.wasm'])
const excludedDirectories = new Set([
  '.git',
  '.worktrees',
  'build',
  'dist',
  'node_modules',
  'out',
  'release',
])

function filesBelow(relative) {
  const absolute = join(root, relative)
  if (!existsSync(absolute)) return []
  return readdirSync(absolute, { withFileTypes: true }).flatMap((entry) => {
    if (entry.isDirectory()) {
      return excludedDirectories.has(entry.name) ? [] : filesBelow(join(relative, entry.name))
    }
    return entry.isFile() ? [join(absolute, entry.name)] : []
  })
}

test('the reviewed optional runtime source packages are present', () => {
  assert.deepEqual(
    runtimePackages.filter((path) => !existsSync(join(root, path))),
    [],
  )
})

test('the base repository does not bundle a Codex executable', () => {
  const bundled = filesBelow('.')
    .filter((path) => binaryExtensions.has(extname(path).toLowerCase()))
    .filter((path) => /(?:codex|app-server)/i.test(basename(path)))
  assert.deepEqual(bundled, [])
})

test('root verification runs the positive optional-runtime contracts', () => {
  const scripts = JSON.parse(readFileSync(join(root, 'package.json'), 'utf8')).scripts
  assert.match(scripts.test, /test:runtime-policy/)
  assert.match(scripts['test:runtime-policy'], /optional-runtime-policy\.test\.mjs/)
  for (const workspace of ['@wiswork/agent-runtime', '@wiswork/codex-bridge']) {
    assert.match(scripts.test, new RegExp(`test -w ${workspace.replace('/', '\\/')}`))
    assert.match(scripts.typecheck, new RegExp(`typecheck -w ${workspace.replace('/', '\\/')}`))
  }
  assert.equal(scripts['test:runtime-removal'], undefined)
})
