import assert from 'node:assert/strict'
import { test } from 'node:test'
import { mkdtempSync, mkdirSync, writeFileSync, readFileSync } from 'node:fs'
import { createRequire } from 'node:module'
import { tmpdir } from 'node:os'
import { join } from 'node:path'
import {
  fingerprint,
  selectBuilds,
  parseArgs,
  sourceRoots,
  verifiedBuildSnapshot,
  buildEnvironment,
  runBoundedCommand,
} from './dogfood-mac.mjs'

test('bounded commands stop only their own spawned process group on timeout', async () => {
  await assert.rejects(
    runBoundedCommand(
      process.execPath,
      ['-e', 'setInterval(() => {}, 1000)'],
      { stdio: 'ignore' },
      100,
    ),
    /timed out/,
  )
})

test('never blesses sources edited during a build with older output', () => {
  const before = { shell: { source: 'old', output: 'old-output' } }
  assert.throws(
    () => verifiedBuildSnapshot(before, { shell: { source: 'new', output: 'built-old-source' } }),
    /changed during build/,
  )
  assert.deepEqual(
    verifiedBuildSnapshot(before, { shell: { source: 'old', output: 'new-output' } }),
    { shell: { source: 'old', output: 'new-output' } },
  )
})
test('dogfood drops externally selected native package inputs', () => {
  assert.equal(
    buildEnvironment({ WISWORK_TECTONIC_SOURCE: '/unverified', PATH: '/bin' })
      .WISWORK_TECTONIC_SOURCE,
    undefined,
  )
})

test('shell bundles editor sources and markdown imports Docs icons', () => {
  assert.ok(sourceRoots('shell').includes('apps/docs'))
  assert.ok(sourceRoots('shell').includes('apps/latex'))
  assert.ok(sourceRoots('shell').includes('tools'))
  assert.deepEqual(sourceRoots('markdown'), ['apps/markdown', 'apps/docs'])
  assert.deepEqual(sourceRoots('docs'), ['apps/docs'])
})

test('packaging cannot inherit release identity or update endpoints', () => {
  process.env.WISWORK_ITERATION_COMMIT = 'abcdef0123456789'
  process.env.WISWORK_ITERATION_BUILT_AT = '2026-09-03T00:00:00Z'
  process.env.WISWORK_UPDATE_PROVIDER = 'generic'
  process.env.WISWORK_UPDATE_URL = 'https://example.com/release'
  process.env.WISWORK_TECTONIC_SOURCE = '/unverified'
  const config = createRequire(import.meta.url)('../apps/shell/electron-builder.dogfood.cjs')
  assert.equal(config.appId, 'com.atominnolab.wiswork.dogfood')
  assert.equal(config.extraMetadata.wisworkIteration.mode, 'dogfood')
  assert.deepEqual(config.protocols, [])
  assert.deepEqual(config.fileAssociations, [])
  assert.equal(config.publish, null)
  assert.equal(config.mac.identity, null)
  assert.equal(config.mac.notarize, false)
  assert.deepEqual(config.mac.target, [{ target: 'dir', arch: [process.arch] }])
  assert.equal(
    config.mac.extraResources.find((entry) => entry.to === 'native/tectonic').from,
    '../latex/native/tectonic',
  )
})
test('startup isolates profile before lifecycle and gates production entry points', () => {
  const source = readFileSync(new URL('../apps/shell/src/main/index.ts', import.meta.url), 'utf8')
  assert.ok(
    source.indexOf(
      "app.setPath('userData', join(app.getPath('appData'), iterationIdentity.productName))",
    ) < source.indexOf('app.requestSingleInstanceLock'),
  )
  assert.match(source, /if \(!iterationIdentity\)\s+registerAuthProtocolRouting/)
  assert.match(source, /if \(app.isPackaged && !iterationIdentity\)\s+migrateLegacyUserData/)
  assert.match(source, /if \(!iterationIdentity\) initAutoUpdater/)
  assert.match(source, /if \(!iterationIdentity\) applyUpdateChannel/)
})

test('strict command options default to launch but allow a read-only plan', () => {
  assert.deepEqual(parseArgs(['--dry-run', '--no-launch']), { dryRun: true, launch: false })
  assert.throws(() => parseArgs(['--production']))
})
test('fingerprints content including untracked files, not timestamps', () => {
  const root = mkdtempSync(join(tmpdir(), 'dogfood-test-'))
  mkdirSync(join(root, 'src'))
  writeFileSync(join(root, 'src/a'), 'one')
  const before = fingerprint(root, ['src'])
  writeFileSync(join(root, 'src/a'), 'two')
  assert.notEqual(fingerprint(root, ['src']), before)
})
test('cache requires both matching source fingerprint and intact output', () => {
  const current = { docs: { source: 'a', output: 'b' }, shell: { source: 'c', output: 'd' } }
  assert.deepEqual(selectBuilds(current, {}), ['docs', 'shell'])
  assert.deepEqual(selectBuilds(current, current), [])
  assert.deepEqual(selectBuilds({ ...current, docs: { source: 'a', output: null } }, current), [
    'docs',
  ])
  assert.deepEqual(
    selectBuilds({ ...current, docs: { source: 'changed', output: 'b' } }, current),
    ['docs'],
  )
})
