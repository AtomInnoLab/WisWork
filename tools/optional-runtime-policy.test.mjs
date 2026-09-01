import assert from 'node:assert/strict'
import { existsSync, mkdirSync, mkdtempSync, readFileSync, rmSync, writeFileSync } from 'node:fs'
import { tmpdir } from 'node:os'
import { join } from 'node:path'
import test from 'node:test'
import { fileURLToPath } from 'node:url'
import { assertOptionalRuntimePackagingPolicy } from './optional-runtime-policy.mjs'

const root = fileURLToPath(new URL('..', import.meta.url))
const runtimePackages = ['packages/agent-runtime', 'packages/codex-bridge']

test('the reviewed optional runtime source packages are present', () => {
  assert.deepEqual(
    runtimePackages.filter((path) => !existsSync(join(root, path, 'package.json'))),
    [],
  )
})

test('the packaging configuration and current output inventory do not bundle Codex', () => {
  assert.doesNotThrow(() => assertOptionalRuntimePackagingPolicy({ root }))
})

test('the policy rejects injected binaries, archives, and extraResources without broad scans', () => {
  const fixture = mkdtempSync(join(tmpdir(), 'wiswork-runtime-policy-'))
  try {
    mkdirSync(join(fixture, 'apps/shell/out'), { recursive: true })
    mkdirSync(join(fixture, 'apps/shell/release'), { recursive: true })
    writeFileSync(
      join(fixture, 'apps/shell/electron-builder.cjs'),
      "module.exports={files:['out/**']}\n",
    )
    for (const name of ['codex', 'codex.exe', 'codex-app-server.zip', 'codex.tar.gz']) {
      writeFileSync(join(fixture, 'apps/shell/out', name), 'fixture')
      assert.throws(() =>
        assertOptionalRuntimePackagingPolicy({
          root: fixture,
          packagingConfig: { files: ['out/**'] },
        }),
      )
      rmSync(join(fixture, 'apps/shell/out', name))
    }
    assert.throws(() =>
      assertOptionalRuntimePackagingPolicy({
        root: fixture,
        packagingConfig: { extraResources: [{ from: 'components/codex-app-server.tar.gz' }] },
      }),
    )
    writeFileSync(join(fixture, 'apps/shell/release/codex-app-server.zip'), 'fixture')
    assert.throws(() => assertOptionalRuntimePackagingPolicy({ root: fixture }))
    rmSync(join(fixture, 'apps/shell/release/codex-app-server.zip'))
    writeFileSync(
      join(fixture, 'apps/shell/electron-builder.cjs'),
      "module.exports={extraResources:[{from:'components/codex'}]}\n",
    )
    assert.throws(() => assertOptionalRuntimePackagingPolicy({ root: fixture }))
  } finally {
    rmSync(fixture, { recursive: true, force: true })
  }
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
