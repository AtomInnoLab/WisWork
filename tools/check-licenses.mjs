/**
 * Dependency license gate: every production npm dependency must
 * carry a license from the permissive allowlist, so a copyleft dependency
 * cannot slip into a release. Reads license fields from package-lock.json
 * (no install needed); the Rust sidecar equivalent is cargo-deny
 * (apps/sheets/native/xlsx-engine/deny.toml).
 */
import { readFileSync } from 'node:fs'
import { createHash } from 'node:crypto'
import { join, dirname } from 'node:path'
import { fileURLToPath } from 'node:url'

const ROOT = join(dirname(fileURLToPath(import.meta.url)), '..')

const ALLOWED = new Set([
  'MIT',
  'MIT-0',
  'Apache-2.0',
  'ISC',
  'BSD-2-Clause',
  'BSD-3-Clause',
  '0BSD',
  'BlueOak-1.0.0',
  'CC0-1.0',
  'CC-BY-4.0',
  'Zlib',
  'Unlicense',
  'Python-2.0',
  'Unicode-3.0',
  'OFL-1.1',
])

/** Packages whose published package.json lacks a license field; license
 * verified manually against the LICENSE file shipped in the package. */
const EXCEPTIONS = {
  '@univerjs/telemetry': 'Apache-2.0',
}

/** Minimal SPDX expression check: OR passes if any branch is allowed,
 * AND requires every branch, WITH falls back to the base license. */
function isAllowed(expr) {
  let s = expr.trim()
  while (s.startsWith('(') && s.endsWith(')')) {
    const inner = s.slice(1, -1)
    let depth = 0
    let balanced = true
    for (const ch of inner) {
      if (ch === '(') depth++
      else if (ch === ')') depth--
      if (depth < 0) balanced = false
    }
    if (!balanced || depth !== 0) break
    s = inner.trim()
  }
  const splitTop = (sep) => {
    const parts = []
    let depth = 0
    let start = 0
    for (let i = 0; i <= s.length - sep.length; i++) {
      if (s[i] === '(') depth++
      else if (s[i] === ')') depth--
      else if (depth === 0 && s.startsWith(sep, i)) {
        parts.push(s.slice(start, i))
        start = i + sep.length
        i += sep.length - 1
      }
    }
    parts.push(s.slice(start))
    return parts
  }
  const orParts = splitTop(' OR ')
  if (orParts.length > 1) return orParts.some(isAllowed)
  const andParts = splitTop(' AND ')
  if (andParts.length > 1) return andParts.every(isAllowed)
  // legacy "A/B" dual-license shorthand
  const slashParts = s.includes('/') ? s.split('/') : [s]
  if (slashParts.length > 1) return slashParts.some(isAllowed)
  const withParts = s.split(' WITH ')
  return ALLOWED.has(withParts[0].trim())
}

const lock = JSON.parse(readFileSync(join(ROOT, 'package-lock.json'), 'utf8'))
const violations = []

for (const [path, info] of Object.entries(lock.packages)) {
  const idx = path.lastIndexOf('node_modules/')
  if (idx === -1) continue // workspace roots
  if (info.dev || info.link) continue
  const name = path.slice(idx + 'node_modules/'.length)
  const license = info.license || EXCEPTIONS[name]
  if (!license) {
    violations.push(`${name}: no license field in lockfile (add to EXCEPTIONS after verifying)`)
  } else if (!isAllowed(license)) {
    violations.push(`${name}: ${license}`)
  }
}

const tectonicManifest = JSON.parse(
  readFileSync(join(ROOT, 'tools/tectonic/manifest.json'), 'utf8'),
)
const tectonicVersion = tectonicManifest?.tectonic?.version
const tectonicLicense = tectonicManifest?.tectonic?.license?.spdx
const tectonicLicenseSource = tectonicManifest?.tectonic?.license?.sourceUrl
const tectonicLicenseText = readFileSync(join(ROOT, 'tools/tectonic/LICENSE'), 'utf8')
const tectonicLicenseHash = createHash('sha256').update(tectonicLicenseText).digest('hex')
if (tectonicVersion !== '0.16.9') violations.push(`tectonic: unexpected version ${tectonicVersion}`)
if (typeof tectonicLicense !== 'string' || !isAllowed(tectonicLicense)) {
  violations.push(`tectonic: ${tectonicLicense ?? 'no SPDX license'}`)
}
try {
  if (new URL(tectonicLicenseSource).protocol !== 'https:')
    violations.push('tectonic: license source must use HTTPS')
} catch {
  violations.push('tectonic: license source is invalid')
}
if (
  tectonicLicenseHash !== '814a258f76e420b25cb3c07172eb2b3956f34cefbf0a650413b78e65c425f306' ||
  !tectonicLicenseText.includes('Tectonic is licensed under the MIT License') ||
  !tectonicLicenseText.includes('THE SOFTWARE IS PROVIDED "AS IS"')
) {
  violations.push('tectonic: vendored MIT license text is incomplete')
}

const codexManifest = JSON.parse(readFileSync(join(ROOT, 'tools/codex/manifest.json'), 'utf8'))
const codexComponent = codexManifest?.component
const codexHashes = new Set(codexComponent?.assets?.map((asset) => asset.sha256))
for (const expected of [
  '552ec417bd679532a531175e705979b9908e575b54ca267f461620b77441a2ad',
  '3124de77fcb1f2eed35e959453ebd49cc4e786cf9236414948f02d58026728f8',
  'c8908d687cf7caa3074921479726db32f96a295372c3544f1e96919a7254951f',
])
  if (!codexHashes.has(expected))
    violations.push(`enhanced-mode component: missing reviewed digest ${expected}`)
if (
  codexComponent?.version !== '0.147.0' ||
  codexComponent?.license?.spdx !== 'Apache-2.0' ||
  !isAllowed(codexComponent?.license?.spdx) ||
  codexComponent?.license?.sourceUrl !==
    'https://github.com/openai/codex/blob/rust-v0.147.0/LICENSE'
) {
  violations.push('enhanced-mode component: invalid pinned license metadata')
}
for (const [path, expectedHash] of [
  ['tools/codex/NOTICE', 'fbe315c31ec234c6eeff5bdd0d06393849ef1f9a1a172a9fa2d1129b8c58b8b7'],
  ['tools/codex/LICENSE-V8', 'b09b68442c92b871a7e4aad1a712302acec44599b961d2e4b2f413836bfd19f7'],
  [
    'tools/codex/LICENSE-RUSTY-V8',
    'e0658d4bc74dc1cf48ae9b442e1493ec530d6fbc4db69efbb5dbfb288e04bdd8',
  ],
])
  if (
    createHash('sha256')
      .update(readFileSync(join(ROOT, path), 'utf8'))
      .digest('hex') !== expectedHash
  ) {
    violations.push(`enhanced-mode component: unreviewed legal text ${path}`)
  }

if (violations.length > 0) {
  console.error('Disallowed or unknown licenses in production dependencies:\n')
  for (const v of violations) console.error(`  ${v}`)
  console.error('\nAllowlist lives in tools/check-licenses.mjs.')
  process.exit(1)
}

console.log('All production npm dependency licenses are within the allowlist.')
