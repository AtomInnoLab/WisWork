import assert from 'node:assert/strict'
import { createHash } from 'node:crypto'
import {
  existsSync,
  mkdirSync,
  mkdtempSync,
  readFileSync,
  rmSync,
  symlinkSync,
  writeFileSync,
} from 'node:fs'
import { createRequire } from 'node:module'
import { tmpdir } from 'node:os'
import { join } from 'node:path'
import test from 'node:test'
import { fileURLToPath } from 'node:url'
import { assertOptionalRuntimePackagingPolicy } from './optional-runtime-policy.mjs'

const root = fileURLToPath(new URL('..', import.meta.url))
const require = createRequire(import.meta.url)
const { createPackageInputs } = require('../apps/shell/packaging-inputs.cjs')
const runtimePackages = ['packages/agent-runtime', 'packages/codex-bridge']

test('the reviewed optional runtime source packages are present', () => {
  assert.deepEqual(
    runtimePackages.filter((path) => !existsSync(join(root, path, 'package.json'))),
    [],
  )
})

test('the packaging configuration and current output inventory do not bundle Codex', () => {
  assert.doesNotThrow(() => assertOptionalRuntimePackagingPolicy({ root }))
  assert.throws(() =>
    assertOptionalRuntimePackagingPolicy({
      root,
      packageInputs: createPackageInputs({ tectonicSource: ['components', 'codex'].join('/') }),
    }),
  )

  const builderPath = require.resolve('../apps/shell/electron-builder.cjs')
  const previousSource = process.env.WISWORK_TECTONIC_SOURCE
  try {
    process.env.WISWORK_TECTONIC_SOURCE = ['components', 'codex'].join('/')
    delete require.cache[builderPath]
    assert.throws(() => assertOptionalRuntimePackagingPolicy({ root }))
  } finally {
    if (previousSource === undefined) delete process.env.WISWORK_TECTONIC_SOURCE
    else process.env.WISWORK_TECTONIC_SOURCE = previousSource
    delete require.cache[builderPath]
  }
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
    mkdirSync(join(fixture, 'node_modules'), { recursive: true })
    assert.throws(() =>
      assertOptionalRuntimePackagingPolicy({
        root: fixture,
        packagingConfig: { extraResources: [{ from: '../../node_modules' }] },
      }),
    )
    writeFileSync(join(fixture, 'apps/shell/release/codex-app-server.zip'), 'fixture')
    assert.throws(() =>
      assertOptionalRuntimePackagingPolicy({
        root: fixture,
        mode: 'post-package',
        artifactDirectories: [join(fixture, 'apps/shell/release')],
      }),
    )
    rmSync(join(fixture, 'apps/shell/release/codex-app-server.zip'))
    writeFileSync(
      join(fixture, 'apps/shell/electron-builder.cjs'),
      "module.exports={extraResources:[{from:'components/codex'}]}\n",
    )
    assert.throws(() => assertOptionalRuntimePackagingPolicy({ root: fixture }))

    const ordinary = join(fixture, 'artifacts/unpacked/resources/bin')
    mkdirSync(ordinary, { recursive: true })
    writeFileSync(join(ordinary, 'codex'), 'fixture')
    assert.throws(() =>
      assertOptionalRuntimePackagingPolicy({
        root: fixture,
        mode: 'post-package',
        artifactDirectories: [join(fixture, 'artifacts')],
      }),
    )
    rmSync(join(ordinary, 'codex'))
    writeFileSync(join(fixture, 'outside'), 'fixture')
    symlinkSync(join(fixture, 'outside'), join(ordinary, 'linked-component'))
    assert.throws(() =>
      assertOptionalRuntimePackagingPolicy({
        root: fixture,
        mode: 'post-package',
        artifactDirectories: [join(fixture, 'artifacts')],
      }),
    )
    rmSync(join(ordinary, 'linked-component'))
    writeFileSync(join(ordinary, 'allowed-component'), 'fixture')
    symlinkSync(join(ordinary, 'allowed-component'), join(ordinary, 'linked-component'))
    assert.doesNotThrow(() =>
      assertOptionalRuntimePackagingPolicy({
        root: fixture,
        mode: 'post-package',
        artifactDirectories: [join(fixture, 'artifacts')],
      }),
    )
    assert.throws(() =>
      assertOptionalRuntimePackagingPolicy({
        root: fixture,
        mode: 'post-package',
        artifactDirectories: [join(fixture, 'missing')],
      }),
    )
  } finally {
    rmSync(fixture, { recursive: true, force: true })
  }
})

test('root verification runs the positive optional-runtime contracts', () => {
  const scripts = JSON.parse(readFileSync(join(root, 'package.json'), 'utf8')).scripts
  assert.match(scripts.test, /test:runtime-policy/)
  assert.match(scripts['test:runtime-policy'], /optional-runtime-policy\.test\.mjs/)
  assert.match(scripts['test:runtime-package'], /--mode post-package/)
  for (const script of ['dist:mac', 'dist:win', 'dist:linux']) {
    assert.match(scripts[script], /test:runtime-package/)
  }
  for (const workspace of ['@wiswork/agent-runtime', '@wiswork/codex-bridge']) {
    assert.match(scripts.test, new RegExp(`test -w ${workspace.replace('/', '\\/')}`))
    assert.match(scripts.typecheck, new RegExp(`typecheck -w ${workspace.replace('/', '\\/')}`))
  }
  assert.equal(scripts['test:runtime-removal'], undefined)
})

test('post-package inventory rejects renamed component bytes by hash', () => {
  const fixture = mkdtempSync(join(tmpdir(), 'wiswork-runtime-policy-hash-'))
  try {
    const artifact = join(fixture, 'release/unpacked/resources/renamed-helper.exe')
    mkdirSync(join(fixture, 'release/unpacked/resources'), { recursive: true })
    const bytes = Buffer.from('small-known-component-fixture')
    writeFileSync(artifact, bytes)
    const digest = createHash('sha256').update(bytes).digest('hex')
    assert.throws(
      () =>
        assertOptionalRuntimePackagingPolicy({
          root: fixture,
          mode: 'post-package',
          artifactDirectories: [join(fixture, 'release')],
          knownHashes: new Set([digest]),
        }),
      /known optional Codex component bytes/,
    )
  } finally {
    rmSync(fixture, { recursive: true, force: true })
  }
})

test('source inventory rejects executable magic outside the native input allowlist', () => {
  const fixture = mkdtempSync(join(tmpdir(), 'wiswork-runtime-policy-magic-'))
  try {
    const out = join(fixture, 'apps/shell/out')
    mkdirSync(out, { recursive: true })
    writeFileSync(join(out, 'renamed-helper.bin'), Buffer.from([0x4d, 0x5a, 0, 0]))
    assert.throws(
      () =>
        assertOptionalRuntimePackagingPolicy({
          root: fixture,
          packagingConfig: { files: ['out/**'] },
          knownHashes: new Set(),
        }),
      /executable package input is not allowlisted/,
    )
    rmSync(join(out, 'renamed-helper.bin'))
    const disguised = join(fixture, 'apps/shell/malware/tectonic')
    mkdirSync(join(fixture, 'apps/shell/malware'), { recursive: true })
    writeFileSync(disguised, Buffer.from([0x7f, 0x45, 0x4c, 0x46]))
    assert.throws(() =>
      assertOptionalRuntimePackagingPolicy({
        root: fixture,
        packagingConfig: { extraResources: [{ from: 'malware/tectonic' }] },
        knownHashes: new Set(),
      }),
    )
    rmSync(disguised)
    const tectonic = join(fixture, 'apps/latex/native/tectonic')
    mkdirSync(join(fixture, 'apps/latex/native'), { recursive: true })
    writeFileSync(tectonic, Buffer.from([0x7f, 0x45, 0x4c, 0x46]))
    assert.doesNotThrow(() =>
      assertOptionalRuntimePackagingPolicy({
        root: fixture,
        packagingConfig: { extraResources: [{ from: '../latex/native/tectonic' }] },
        knownHashes: new Set(),
      }),
    )
  } finally {
    rmSync(fixture, { recursive: true, force: true })
  }
})
