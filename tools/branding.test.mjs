import { execFileSync } from 'node:child_process'
import { createHash } from 'node:crypto'
import assert from 'node:assert/strict'
import { existsSync, readFileSync } from 'node:fs'
import { createRequire } from 'node:module'
import { join } from 'node:path'
import { afterEach, test } from 'node:test'

const require = createRequire(import.meta.url)
const root = join(import.meta.dirname, '..')

function json(path) {
  return JSON.parse(readFileSync(join(root, path), 'utf8'))
}

afterEach(() => {
  delete process.env.WISWORK_UPDATE_URL
  delete process.env.WISWORK_UNSIGNED_MAC_BUILD
  delete process.env.WISWORK_TECTONIC_SOURCE
  delete process.env.GENOFFICE_UPDATE_URL
  delete require.cache[require.resolve('../apps/shell/electron-builder.cjs')]
})

test('package-lock workspace identities use only the WisWork scope and root name', () => {
  const lock = json('package-lock.json')
  assert.equal(lock.packages[''].name, 'wiswork')
  for (const [path, metadata] of Object.entries(lock.packages)) {
    if (!/^(apps|packages)\//.test(path) || path.includes('/node_modules/')) continue
    assert.match(metadata.name, /^@wiswork\//, path)
    assert.doesNotMatch(metadata.name, /genoffice|genspark|com\.genoffice/i, path)
  }
})

test('all distributable apps use WisWork names and AtomInnoLab bundle identifiers', () => {
  const expected = new Map([
    ['apps/shell/package.json', ['WisWork', undefined]],
    ['apps/docs/package.json', ['WisWork Docs', 'com.atominnolab.wiswork.docs']],
    ['apps/sheets/package.json', ['WisWork Sheets', 'com.atominnolab.wiswork.sheets']],
    ['apps/slides/package.json', ['WisWork Slides', 'com.atominnolab.wiswork.slides']],
    ['apps/pdf/package.json', ['WisWork PDF', 'com.atominnolab.wiswork.pdf']],
    ['apps/latex/package.json', ['WisWork LaTeX', 'com.atominnolab.wiswork.latex']],
  ])
  for (const [path, [productName, appId]] of expected) {
    const pkg = json(path)
    assert.equal(pkg.productName, productName, path)
    assert.equal(pkg.author, 'AtomInnoLab', path)
    if (appId && pkg.build) assert.equal(pkg.build.appId, appId, path)
  }
})

test('shell packaging uses WisWork update URL and exact product metadata', () => {
  process.env.WISWORK_UPDATE_URL = 'https://updates.example/wiswork/'
  const config = require('../apps/shell/electron-builder.cjs')
  assert.equal(config.appId, 'com.atominnolab.wiswork')
  assert.equal(config.productName, 'WisWork')
  assert.equal(config.artifactName, 'WisWork-${version}-${arch}.${ext}')
  assert.equal(config.publish[0].url, 'https://updates.example/wiswork')
  assert.equal(config.mac.identity, undefined)
  assert.equal(config.mac.notarize, true)
  assert.equal(config.dmg.sign, true)
})

test('legacy update URL remains an explicit one-release compatibility fallback', () => {
  process.env.GENOFFICE_UPDATE_URL = 'https://legacy.example/channel/'
  const config = require('../apps/shell/electron-builder.cjs')
  assert.equal(config.publish[0].url, 'https://legacy.example/channel')
})

test('unsigned macOS test packaging disables signing and notarization only when explicitly enabled', () => {
  process.env.WISWORK_UNSIGNED_MAC_BUILD = '1'
  const config = require('../apps/shell/electron-builder.cjs')
  assert.equal(config.mac.identity, null)
  assert.equal(config.mac.notarize, false)
  assert.equal(config.dmg.sign, false)
  assert.equal(config.afterAllArtifactBuild, undefined)
})

test('macOS packaging workflow builds an arm64 sidecar and uploads dmg and zip artifacts', () => {
  const workflow = readFileSync(join(root, '.github/workflows/package-macos.yml'), 'utf8')
  assert.match(workflow, /aarch64-apple-darwin/)
  assert.match(workflow, /WISWORK_UNSIGNED_MAC_BUILD:\s*['"]1['"]/)
  assert.match(workflow, /electron-builder --config electron-builder\.cjs --mac dmg zip --arm64/)
  assert.match(workflow, /release\/\*\.dmg/)
  assert.match(workflow, /release\/\*\.zip/)
})

test('shell packages the LaTeX renderer and only the verified Tectonic executable', () => {
  process.env.WISWORK_TECTONIC_SOURCE = '/tmp/verified-tectonic'
  const config = require('../apps/shell/electron-builder.cjs')
  assert.ok(
    config.extraResources.some(
      (entry) => entry.from === '../latex/out' && entry.to === 'modules/latex',
    ),
  )
  assert.ok(
    config.extraResources.some(
      (entry) =>
        entry.from === '../../tools/tectonic/manifest.json' &&
        entry.to === 'native/tectonic-manifest.json',
    ),
  )
  assert.ok(
    config.mac.extraResources.some(
      (entry) => entry.from === '/tmp/verified-tectonic' && entry.to === 'native/tectonic',
    ),
  )
  assert.equal(JSON.stringify(config).includes('tectonic-default-bundle-v33'), false)
})

test('macOS workflow fetches, verifies, injects, and inspects the arm64 Tectonic sidecar', () => {
  const workflow = readFileSync(join(root, '.github/workflows/package-macos.yml'), 'utf8')
  for (const path of [
    'apps/shell/electron-builder.cjs',
    'apps/latex/**',
    'packages/latex-project/**',
    'packages/latex-compiler/**',
    'packages/pdf-viewer/**',
    'tools/tectonic/**',
    'package-lock.json',
  ]) {
    assert.ok(workflow.includes(`- '${path}'`) || workflow.includes(`- "${path}"`), path)
  }
  assert.match(workflow, /node tools\/fetch-tectonic\.mjs --platform darwin-arm64 --output [^\n]+/)
  assert.match(workflow, /file [^\n]*tectonic[^\n]*\| grep -q arm64/)
  assert.match(workflow, /tectonic[^\n]*--version[^\n]*0\.16\.9/)
  assert.match(workflow, /WISWORK_TECTONIC_SOURCE:/)
  assert.match(workflow, /modules\/latex\/renderer\/index\.html/)
  assert.match(workflow, /Contents\/Resources\/native\/tectonic/)
  assert.match(workflow, /SHA256SUMS\.txt/)
})

test('generated notices and developer docs cover LaTeX and pinned Tectonic metadata', () => {
  // The notices file is a git-ignored build artifact that only `npm run notices`
  // (a dist:* step) produces, so it is absent on a fresh checkout — the branding
  // gate runs straight after `npm ci`. Generate it here instead of ordering a
  // packaging step ahead of the gate.
  const noticesPath = join(root, 'apps/shell/build/THIRD-PARTY-NOTICES.txt')
  if (!existsSync(noticesPath)) {
    execFileSync(process.execPath, [join(root, 'tools/gen-third-party-notices.mjs')], {
      cwd: root,
      stdio: ['ignore', 'ignore', 'inherit'],
    })
  }
  const notices = readFileSync(noticesPath, 'utf8')
  assert.match(notices, /codemirror-lang-latex v0\.4\.2 — MIT/)
  assert.match(notices, /Tectonic 0\.16\.9 — MIT/)
  assert.match(
    notices,
    /Tectonic 0\.16\.9 — MIT[\s\S]*?Tectonic is licensed under the MIT License[\s\S]*?THE SOFTWARE IS PROVIDED "AS IS"/,
  )
  const docs = readFileSync(join(root, 'docs/development/latex.md'), 'utf8')
  assert.match(docs, /first compile/i)
  assert.match(docs, /offline/i)
  assert.match(docs, /cache/i)
  assert.match(docs, /2\.8\s*(?:GB|GiB)/i)
  assert.doesNotMatch(docs, /sk-[a-z0-9]{20,}|1234567890/i)
  assert.match(readFileSync(join(root, 'README.md'), 'utf8'), /docs\/development\/latex\.md/)
})

test('LaTeX polls authoritative bundle state even for AI-triggered compilation', () => {
  const source = readFileSync(join(root, 'apps/latex/src/renderer/App.tsx'), 'utf8')
  const effect = source.slice(
    source.indexOf('window.latexApi.getBundleStatus'),
    source.indexOf('window.latexApi.onEditFlushRequest'),
  )
  assert.match(effect, /window\.setInterval\(\(\) => void update\(\), 250\)/)
  assert.doesNotMatch(effect, /if \(!compiling\)/)
})

test('CI runs the branding gate immediately after dependency installation', () => {
  const workflow = readFileSync(join(root, '.github/workflows/ci.yml'), 'utf8')
  const install = workflow.indexOf('run: npm ci')
  const branding = workflow.indexOf('run: npm run test:branding')
  assert.ok(install >= 0)
  assert.ok(branding > install)
  assert.ok(branding < workflow.indexOf('run: npm run licenses', install))
})

test('NOTICE preserves the complete upstream attribution block verbatim', () => {
  const notice = readFileSync(join(root, 'NOTICE'), 'utf8')
  assert.match(
    notice,
    /GenOffice\nCopyright 2026 Mainfunc, Inc\.\nThis product includes software developed at Mainfunc, Inc\./,
  )
})

test('active repository text contains no unallowlisted legacy product branding', () => {
  const files = execFileSync('git', ['ls-files', '-co', '--exclude-standard'], {
    cwd: root,
    encoding: 'utf8',
  })
    .trim()
    .split('\n')
    .filter(Boolean)
  const legacy = /genoffice|genspark|genteam|com\.genoffice/i
  const allowedLegacyLine = (path, line) => {
    if (path === 'apps/shell/src/main/user-data-migration.ts')
      return line.trim() === "const LEGACY_PRODUCT_NAMES = ['GenOffice', 'AI Office'] as const"
    if (path === 'apps/shell/tests/user-data-migration.test.ts')
      return /(?:'GenOffice'|'AI Office'|copies GenOffice|falls back to AI Office)/.test(line)
    if (path === 'apps/shell/tests/removed-genteam.test.ts')
      return /removed GenTeam|openGenTeam|home:open-genteam/.test(line)
    if (path === 'apps/shell/electron-builder.cjs') return line.includes('GENOFFICE_UPDATE_URL')
    if (path === 'NOTICE')
      return ['GenOffice', 'This product includes software developed at Mainfunc, Inc.'].includes(
        line.trim(),
      )
    if (path === 'ee/LICENSE') return line.trim() === 'GenOffice Enterprise License'
    return false
  }
  const violations = []
  for (const path of files) {
    if (path === 'tools/branding.test.mjs' || path.startsWith('docs/superpowers/')) continue
    let source
    try {
      source = readFileSync(join(root, path), 'utf8')
    } catch {
      continue
    }
    if (source.includes('\0')) continue
    source.split('\n').forEach((line, index) => {
      if (!legacy.test(line) && !line.includes('AI Office')) return
      if (allowedLegacyLine(path, line)) return
      violations.push(`${path}:${index + 1}: ${line.trim()}`)
    })
  }
  assert.deepEqual(violations, [])
})

test('all shipped bitmap and platform icons have been replaced with WisWork artwork', () => {
  const legacyHashes = new Set([
    'cd04254b02d9f8921f0d27f6abc1f42903b9ce70621854ae8df2ee603a90ca84',
    'acdd167f36015b98f45bc2cd31f0ac88fd57920e171d35b705fd831be95055b7',
    '3cf594e4aebcc24f38d64b895be6b71a5fd1dcea8df1d61948ba6cbda21b49db',
    '5dbcc7c50e29639eaea4a336b70e25466943abc09cc320b1d934e6b006da13a0',
  ])
  const pngs = [
    'apps/shell/build/icon.png',
    'apps/shell/build/icon-mac.png',
    'apps/shell/src/renderer/src/assets/app-icon.png',
    'apps/docs/src/renderer/assets/app-icon.png',
    'apps/sheets/src/renderer/assets/app-icon.png',
    'apps/slides/src/renderer/assets/app-icon.png',
  ]
  for (const path of pngs) {
    const bytes = readFileSync(join(root, path))
    assert.equal(bytes.subarray(1, 4).toString(), 'PNG', path)
    assert.equal(bytes.readUInt32BE(16), 1024, path)
    assert.equal(bytes.readUInt32BE(20), 1024, path)
    assert.equal(legacyHashes.has(createHash('sha256').update(bytes).digest('hex')), false, path)
  }
  for (const path of ['apps/shell/build/icon.icns', 'apps/shell/build/icon.ico']) {
    const bytes = readFileSync(join(root, path))
    assert.equal(legacyHashes.has(createHash('sha256').update(bytes).digest('hex')), false, path)
  }
  assert.equal(readFileSync(join(root, 'apps/shell/build/icon.icns'), 'ascii').slice(0, 4), 'icns')
  assert.deepEqual(
    [...readFileSync(join(root, 'apps/shell/build/icon.ico')).subarray(0, 4)],
    [0, 0, 1, 0],
  )
})
