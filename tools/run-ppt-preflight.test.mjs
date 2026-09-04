import assert from 'node:assert/strict'
import { mkdirSync, mkdtempSync, readFileSync, writeFileSync } from 'node:fs'
import { tmpdir } from 'node:os'
import { dirname, join } from 'node:path'
import test from 'node:test'
import {
  createPptPreflightSteps,
  orchestratePptPreflight,
  resolvePreflightExecutable,
} from './run-ppt-preflight.mjs'

const fakePptx = () => Buffer.concat([Buffer.from('PK'), Buffer.alloc(1_100, 1)])

test('requires live WisUsage credentials and a real Codex executable', () => {
  assert.throws(
    () =>
      createPptPreflightSteps({
        env: {},
        platform: 'darwin',
        executable: '/missing',
        artifactRoot: '/tmp/x',
      }),
    /ppt_preflight_wisusage_token_required/,
  )
  assert.throws(
    () =>
      createPptPreflightSteps({
        env: { WISWORK_REAL_WISUSAGE_TOKEN: 'secret' },
        platform: 'darwin',
        executable: '/missing',
        artifactRoot: '/tmp/x',
      }),
    /ppt_preflight_codex_component_required/,
  )
})

test('finds the installed macOS Enhanced component without packaging', () => {
  const home = mkdtempSync(join(tmpdir(), 'wiswork-preflight-home-'))
  const executable = join(
    home,
    'Library/Application Support/WisWork/components/enhanced-mode/0.147.0/darwin-arm64/bin/codex-app-server',
  )
  mkdirSync(dirname(executable), { recursive: true })
  writeFileSync(executable, 'binary')
  assert.equal(
    resolvePreflightExecutable({ env: {}, platform: 'darwin', arch: 'arm64', home }),
    executable,
  )
})

test('runs contracts, both live modes and source Electron before writing a bounded report', async () => {
  const root = mkdtempSync(join(tmpdir(), 'wiswork-ppt-preflight-'))
  const executable = join(root, 'codex')
  writeFileSync(executable, 'binary')
  const calls = []
  const code = await orchestratePptPreflight({
    env: {
      WISWORK_REAL_WISUSAGE_TOKEN: 'secret',
      WISWORK_CODEX_INTEGRATION_EXECUTABLE: executable,
    },
    platform: 'darwin',
    arch: 'arm64',
    home: root,
    root,
    runStep: async (step) => {
      calls.push(step)
      if (step.name === 'standard-live-ppt')
        writeFileSync(step.extraEnv.WISWORK_STANDARD_PPT_E2E_OUTPUT, fakePptx())
      if (step.name === 'enhanced-live-ppt')
        writeFileSync(step.extraEnv.WISWORK_ENHANCED_PPT_E2E_OUTPUT, fakePptx())
      return 0
    },
  })
  assert.equal(code, 0)
  assert.deepEqual(
    calls.map((call) => call.name),
    ['production-contracts', 'standard-live-ppt', 'enhanced-live-ppt', 'electron-source-e2e'],
  )
  assert.equal(calls[1].extraEnv.WISWORK_REAL_WISUSAGE_TOKEN, 'secret')
  const report = JSON.parse(
    readFileSync(join(root, 'test-results/ppt-preflight/report.json'), 'utf8'),
  )
  assert.equal(report.schema, 'wiswork-ppt-preflight/v1')
  assert.equal(report.artifacts.length, 2)
  assert.equal(JSON.stringify(report).includes('secret'), false)
})

test('stops immediately on failure and emits a failure-only diagnostic report', async () => {
  const root = mkdtempSync(join(tmpdir(), 'wiswork-ppt-preflight-fail-'))
  const executable = join(root, 'codex')
  writeFileSync(executable, 'binary')
  const calls = []
  const code = await orchestratePptPreflight({
    env: {
      WISWORK_REAL_WISUSAGE_TOKEN: 'secret',
      WISWORK_CODEX_INTEGRATION_EXECUTABLE: executable,
    },
    platform: 'darwin',
    arch: 'arm64',
    home: root,
    root,
    runStep: async ({ name }) => {
      calls.push(name)
      return name === 'standard-live-ppt' ? 7 : 0
    },
  })
  assert.equal(code, 7)
  assert.deepEqual(calls, ['production-contracts', 'standard-live-ppt'])
  const report = JSON.parse(
    readFileSync(join(root, 'test-results/ppt-preflight/report.json'), 'utf8'),
  )
  assert.equal(report.status, 'failed')
  assert.equal(report.failedStep, 'standard-live-ppt')
  assert.equal(JSON.stringify(report).includes('secret'), false)
})
