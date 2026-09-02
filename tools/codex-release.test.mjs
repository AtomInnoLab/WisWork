import assert from 'node:assert/strict'
import { existsSync, readFileSync, readdirSync, statSync } from 'node:fs'
import { dirname, join } from 'node:path'
import { fileURLToPath } from 'node:url'
import test from 'node:test'
import { optionalRuntimeKnownHashes } from './optional-runtime-policy.mjs'

const root = join(dirname(fileURLToPath(import.meta.url)), '..')
const read = (path) => readFileSync(join(root, path), 'utf8')
const json = (path) => JSON.parse(read(path))

test('manifest pins only the three official 0.147.0 app-server package assets', () => {
  const manifest = json('tools/codex/manifest.json')
  assert.equal(manifest.component.version, '0.147.0')
  assert.deepEqual(
    manifest.component.assets.map((asset) => [
      asset.platform,
      asset.arch,
      asset.bytes,
      asset.sha256,
    ]),
    [
      [
        'darwin',
        'arm64',
        92265540,
        '552ec417bd679532a531175e705979b9908e575b54ca267f461620b77441a2ad',
      ],
      [
        'darwin',
        'x64',
        99846280,
        '3124de77fcb1f2eed35e959453ebd49cc4e786cf9236414948f02d58026728f8',
      ],
      [
        'win32',
        'x64',
        110054928,
        'c8908d687cf7caa3074921479726db32f96a295372c3544f1e96919a7254951f',
      ],
    ],
  )
  for (const asset of manifest.component.assets) {
    assert.match(
      asset.primaryUrl,
      /^https:\/\/downloads\.wiswork\.com\/components\/codex\/0\.147\.0\//,
    )
    assert.match(
      asset.fallbackUrl,
      /^https:\/\/github\.com\/openai\/codex\/releases\/download\/rust-v0\.147\.0\//,
    )
    assert.equal(
      new URL(asset.primaryUrl).pathname.split('/').at(-1),
      new URL(asset.fallbackUrl).pathname.split('/').at(-1),
    )
    assert.match(asset.layout.entrypoint, /^bin\/codex-app-server(?:\.exe)?$/)
    assert.deepEqual(
      asset.layout.files.filter((file) => file.install).map((file) => file.path),
      [
        asset.layout.entrypoint,
        asset.platform === 'win32' ? 'bin/codex-code-mode-host.exe' : 'bin/codex-code-mode-host',
        'codex-package.json',
      ],
    )
    if (asset.platform === 'darwin') {
      assert.equal(asset.trust.policy, 'macos')
      assert.equal(asset.trust.teamIdentifier, '2DC432GLL2')
      assert.equal(asset.trust.requireNotarization, false)
    }
  }
  assert.deepEqual(
    [...optionalRuntimeKnownHashes(root)].sort(),
    manifest.component.assets
      .flatMap((asset) => [
        asset.sha256,
        ...asset.layout.files.filter((file) => file.install).map((file) => file.sha256),
      ])
      .sort(),
  )
})

test('reviewed repository assets contain no downloaded component binary or archive', () => {
  for (const name of readdirSync(join(root, 'tools/codex'))) {
    const path = join(root, 'tools/codex', name)
    assert.ok(statSync(path).size < 1_000_000, `binary-sized reviewed asset: ${name}`)
    assert.doesNotMatch(name, /\.(?:tar\.gz|zip|exe|dmg)$/i)
  }
})

test('every desktop builder invocation immediately runs the shared post-package checker', () => {
  for (const path of ['.github/workflows/desktop-release.yml']) {
    const workflow = read(path)
    const builderCount = [...workflow.matchAll(/run: npx electron-builder[^\n]*/g)].length
    const checkerCount = [
      ...workflow.matchAll(
        /run: node \.\.\/\.\.\/tools\/optional-runtime-policy\.mjs --mode post-package --artifact-dir release/g,
      ),
    ].length
    const componentGateCount = [
      ...workflow.matchAll(/run: npx tsx tools\/install-enhanced-component\.ts --cache/g),
    ].length
    assert.ok(builderCount > 0)
    assert.equal(checkerCount, builderCount, `${path} must gate every builder invocation`)
    assert.equal(
      componentGateCount,
      builderCount,
      `${path} must verify the real component per build`,
    )
    for (const match of workflow.matchAll(/run: npx electron-builder[^\n]*/g)) {
      const tail = workflow.slice(match.index + match[0].length)
      assert.ok(tail.indexOf('optional-runtime-policy.mjs --mode post-package') >= 0)
      assert.ok(
        tail.indexOf('optional-runtime-policy.mjs --mode post-package') <
          tail.indexOf('uses: actions/upload-artifact'),
      )
    }
    assert.match(workflow, /grep '\^TeamIdentifier=2DC432GLL2\$'/)
  }
  const desktop = read('.github/workflows/desktop-release.yml')
  assert.equal(existsSync(join(root, '.github/workflows/package-macos.yml')), false)
  assert.match(desktop, /Get-AuthenticodeSignature/)
  assert.match(desktop, /SignerCertificate\.Subject -ne 'CN=OpenAI, L\.L\.C\.'/)
})

test('license and notice generators cover the optional app-server', () => {
  assert.match(read('tools/check-licenses.mjs'), /enhanced-mode component: missing reviewed digest/)
  const generator = read('tools/gen-third-party-notices.mjs')
  assert.match(generator, /OpenAI Codex app-server/)
  assert.match(generator, /not included in the base installer/)
})

test('base package source policy still passes without a component cache', () => {
  const policy = read('tools/optional-runtime-policy.mjs')
  assert.match(policy, /optional Codex artifact must not be bundled/)
  assert.match(policy, /post-package mode requires artifact directories/)
})
