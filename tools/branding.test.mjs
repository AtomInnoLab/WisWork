import { execFileSync } from 'node:child_process'
import { createHash } from 'node:crypto'
import assert from 'node:assert/strict'
import { readFileSync } from 'node:fs'
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
})

test('legacy update URL remains an explicit one-release compatibility fallback', () => {
  process.env.GENOFFICE_UPDATE_URL = 'https://legacy.example/channel/'
  const config = require('../apps/shell/electron-builder.cjs')
  assert.equal(config.publish[0].url, 'https://legacy.example/channel')
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
