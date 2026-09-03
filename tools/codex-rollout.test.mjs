import assert from 'node:assert/strict'
import { existsSync, readFileSync } from 'node:fs'
import { dirname, join } from 'node:path'
import { fileURLToPath } from 'node:url'
import test from 'node:test'
import * as yaml from 'js-yaml'
import { sevenHostGoldens } from './enhanced-seven-host-golden.mjs'

const root = join(dirname(fileURLToPath(import.meta.url)), '..')
const read = (path) => readFileSync(join(root, path), 'utf8')

test('seven-host catalog executes production runtime, adapter, transaction and raw proposal paths', () => {
  assert.deepEqual(
    sevenHostGoldens.map(({ host }) => host),
    ['latex', 'slides', 'docs', 'sheets', 'office-word', 'office-excel', 'office-powerpoint'],
  )
  for (const golden of sevenHostGoldens)
    for (const file of golden.files) assert.ok(existsSync(join(root, file)), file)
  const runner = read('tools/enhanced-seven-host-golden.mjs')
  assert.match(runner, /mkdtempSync/)
  assert.match(runner, /Object\.keys\(report\)\.sort\(\)/)
  assert.match(runner, /rmSync\(reportRoot/)
  assert.doesNotMatch(runner, /disableConsoleIntercept/)
  const office = sevenHostGoldens.filter(({ host }) => host.startsWith('office-'))
  assert.equal(new Set(office.map(({ testName }) => testName)).size, 3)
  assert.ok(office.every(({ testName }) => typeof testName === 'string'))
  assert.equal(sevenHostGoldens.find(({ host }) => host === 'slides').commands, undefined)
})

test('native release matrix gates pinned lifecycle, identity, real protocol and base-package absence', () => {
  const workflow = yaml.load(read('.github/workflows/desktop-release.yml'))
  const build = workflow.jobs.build
  assert.deepEqual(build.strategy.matrix.include, [
    {
      id: 'macos-arm64',
      runner: 'macos-14',
      tectonic_platform: 'darwin-arm64',
      cargo_target: 'aarch64-apple-darwin',
      file_arch: 'arm64',
      electron_args: '--mac dmg zip --arm64',
    },
  ])
  assert.equal(workflow.jobs['verify-release'].steps[3].with.pattern, 'macos-arm64')
  const runs = build.steps.flatMap((step) => (typeof step.run === 'string' ? [step.run] : []))
  const commands = runs.join('\n')
  for (const gate of [
    'npm ci',
    'test:codex-rollout',
    'component-manager.test.ts',
    'process-manager.test.ts',
    'codex-runtime.test.ts',
    'elevated-office-program.test.ts',
    'run-codex-native-lifecycle.ts',
    'install-enhanced-component',
    'codex-engine.integration.test',
    'generates and verifies a three-page onboarding deck',
    'TeamIdentifier=2DC432GLL2',
    'optional-runtime-policy.mjs --mode post-package',
  ])
    assert.ok(commands.includes(gate), gate)
  const uploadIndex = build.steps.findIndex(
    (step) => typeof step.uses === 'string' && step.uses.startsWith('actions/upload-artifact@'),
  )
  const absenceIndex = build.steps.findIndex(
    (step) =>
      typeof step.run === 'string' &&
      step.run.includes('optional-runtime-policy.mjs --mode post-package'),
  )
  assert.ok(absenceIndex >= 0 && uploadIndex > absenceIndex)
  const nativeLifecycle = build.steps.find(
    (step) =>
      step.name ===
      'Exercise native component update, restart, crash recovery, kill switches and removal',
  )
  assert.equal(nativeLifecycle?.shell, 'bash')
  assert.match(nativeLifecycle?.run ?? '', /^npx tsx tools\/run-codex-native-lifecycle\.ts /)
})

test('CI executes the real Slides Electron/Konva render acceptance gate', () => {
  const workflow = yaml.load(read('.github/workflows/ci.yml'))
  const e2e = workflow.jobs.e2e
  assert.ok(
    e2e.steps.some(
      (step) =>
        typeof step.run === 'string' &&
        step.run.includes(
          'xvfb-run --auto-servernum -- npm run test:e2e -- e2e/slides-acceptance-render.spec.ts',
        ),
    ),
  )
  assert.ok(
    e2e.steps.some(
      (step) =>
        typeof step.uses === 'string' &&
        step.uses.startsWith('actions/upload-artifact@') &&
        step.if === 'failure()',
    ),
  )
  assert.match(read('package.json'), /"test:e2e": "node tools\/run-slides-acceptance-e2e\.mjs"/)
})

test('rollout and incident runbook names every independent rollback control', () => {
  const runbook = read('docs/development/codex-enhanced-release.md')
  for (const phrase of [
    'raw Office',
    'per-host',
    'global Enhanced',
    'component-version',
    'Standard',
    'applied_unverified',
    'write_pending_quarantined',
  ])
    assert.match(runbook, new RegExp(phrase, 'i'))
})

test('all editor surfaces obtain visible AI state from their production locale adapter', () => {
  for (const path of [
    'apps/latex/src/renderer/ai/AiPanel.tsx',
    'apps/slides/src/renderer/ai/AiPanel.tsx',
    'apps/docs/src/renderer/ai/AiPanel.tsx',
    'apps/sheets/src/renderer/ai/AiChatPanel.tsx',
  ]) {
    const source = read(path)
    assert.match(source, /use(?:I18n|LatexLocale)/)
  }
  const office = read('apps/office-addin/src/App.tsx')
  assert.match(office, /normalizeLang\(locale\)/)
  assert.match(office, /translatePresentationVerification/)
})
