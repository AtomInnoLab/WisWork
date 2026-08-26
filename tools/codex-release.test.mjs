import assert from 'node:assert/strict'
import { existsSync, readFileSync, readdirSync, statSync } from 'node:fs'
import { execFileSync } from 'node:child_process'
import { dirname, join } from 'node:path'
import { fileURLToPath } from 'node:url'
import test from 'node:test'

const root = join(dirname(fileURLToPath(import.meta.url)), '..')
const readJson = (path) => JSON.parse(readFileSync(join(root, path), 'utf8'))

test('optional component manifest is the reviewed darwin-arm64 0.147.0 release', () => {
  const manifest = readJson('tools/codex/manifest.json')
  assert.equal(manifest.component.version, '0.147.0')
  assert.equal(manifest.component.assets.length, 1)
  const asset = manifest.component.assets[0]
  assert.deepEqual([asset.platform, asset.arch], ['darwin', 'arm64'])
  assert.equal(asset.bytes, 107229164)
  assert.equal(asset.sha256, '17b2984eb22b607e3d0c25728252fc90f510e476bad39a6d9f45cdb1aa685432')
  assert.deepEqual(
    asset.layout.files.filter((file) => file.install).map((file) => file.path),
    ['bin/codex', 'bin/codex-code-mode-host'],
  )
})

test('eval and security fixtures meet the approved pilot thresholds', () => {
  const evaluation = readJson('packages/codex-bridge/evals/latex-runtime-pilot.json')
  const cases = evaluation.cases
  assert.ok(cases.length > 0)
  for (const item of cases) {
    assert.equal(item.legacy, 'pass')
    assert.equal('enhanced' in item, false)
    assert.ok(existsSync(join(root, item.automatedEvidence)))
  }

  const security = readJson('packages/codex-bridge/evals/security-matrix.json')
  assert.equal(security.requiredPassRate, 1)
  for (const item of security.cases) {
    assert.equal('status' in item, false)
    assert.ok(existsSync(join(root, item.evidence)))
  }
  for (const required of [
    'auth',
    'token',
    'logs',
    'ipc',
    'path',
    'shell',
    'network',
    'lifecycle',
  ]) {
    assert.ok(
      security.cases.some((item) => item.id.includes(required)),
      `missing ${required}`,
    )
  }

  const protocol = readJson('packages/codex-bridge/evals/protocol-compatibility.json')
  const generated = readJson('packages/codex-bridge/src/generated/schema-manifest.json')
  assert.equal(protocol.runtimeVersion, generated.version)
  assert.equal(protocol.bindingsSha256, generated.bindingsSha256)
  assert.equal(protocol.schemaSha256, generated.sha256['codex_app_server_protocol.schemas.json'])
  assert.equal(protocol.maximumDocumentCallsPerTurn, 1)
})

test('release builds ignore arbitrary executable overrides and base packages contain no component', () => {
  const main = readFileSync(join(root, 'apps/shell/src/main/index.ts'), 'utf8')
  assert.match(main, /!app\.isPackaged\s*&&\s*typeof process\.env\.WISWORK_CODEX_PATH === 'string'/)
  const builder = readFileSync(join(root, 'apps/shell/electron-builder.cjs'), 'utf8')
  assert.doesNotMatch(builder, /codex|enhanced-mode/i)
  const reviewedAssets = readdirSync(join(root, 'tools/codex')).map((name) => `tools/codex/${name}`)
  for (const path of reviewedAssets) {
    assert.ok(statSync(join(root, path)).size < 1_000_000, `unexpected binary-sized asset: ${path}`)
  }
})

test('user-facing copy uses only Standard mode and Enhanced mode naming', () => {
  for (const path of [
    'apps/shell/src/renderer/src/Home.tsx',
    'apps/shell/src/renderer/src/enhanced-mode-view.ts',
    'apps/latex/src/renderer/ai/AiPanel.tsx',
  ]) {
    const source = readFileSync(join(root, path), 'utf8')
    assert.doesNotMatch(source, /["'`]([^"'`]*\bCodex\b[^"'`]*)["'`]/)
  }
  const runtime = readFileSync(join(root, 'apps/shell/src/main/codex-runtime.ts'), 'utf8')
  assert.doesNotMatch(runtime, /Codex stopped\./)
  const shellMain = readFileSync(join(root, 'apps/shell/src/main/index.ts'), 'utf8')
  assert.doesNotMatch(shellMain, /console\.(?:warn|error|log)\([^\n]*codex/i)

  const release = readFileSync(join(root, 'docs/releases/enhanced-mode-pilot.md'), 'utf8')
  const internalHeading = release.indexOf('## Internal implementation and third-party attribution')
  assert.ok(internalHeading > 0)
  assert.doesNotMatch(release.slice(0, internalHeading), /\bCodex\b/i)
  assert.match(release.slice(internalHeading), /OpenAI Codex CLI 0\.147\.0/)
})

test('generated legal notices cover the optional implementation and embedded runtime host', () => {
  const noticesPath = join(root, 'apps/shell/build/THIRD-PARTY-NOTICES.txt')
  execFileSync(process.execPath, [join(root, 'tools/gen-third-party-notices.mjs')], {
    cwd: root,
    stdio: ['ignore', 'ignore', 'inherit'],
  })
  assert.ok(existsSync(noticesPath))
  const notices = readFileSync(noticesPath, 'utf8')
  assert.match(notices, /OpenAI Codex CLI 0\.147\.0 — Apache-2\.0/)
  assert.match(notices, /Copyright 2025 OpenAI/)
  assert.match(notices, /Copyright 2014, the V8 project authors/)
  assert.match(notices, /Copyright \(c\) 2018-2019 the Deno authors/)
})

test('macOS CI keeps base packaging independent and fail-closes the optional pilot', () => {
  const workflow = readFileSync(join(root, '.github/workflows/package-macos.yml'), 'utf8')
  assert.match(workflow, /enhanced-mode-arm64:/)
  assert.match(workflow, /install-enhanced-component\.ts --cache/)
  assert.match(workflow, /codesign --verify --strict --verbose=4/)
  assert.match(workflow, /spctl --assess --type execute --verbose=4/)
  assert.match(workflow, /WISWORK_CODEX_INTEGRATION_EXECUTABLE:/)
  assert.match(workflow, /tests\/app-server\.integration\.test\.ts/)
  assert.match(workflow, /npm run eval:enhanced-mode/)
  assert.match(workflow, /Verify base package excludes optional Enhanced mode binaries/)
})
