import assert from 'node:assert/strict'
import { createHash } from 'node:crypto'
import { mkdtemp, mkdir, writeFile, rm } from 'node:fs/promises'
import { tmpdir } from 'node:os'
import { resolve } from 'node:path'
import { test } from 'node:test'
import {
  inspectOfficeBuild,
  inspectRelayHealth,
  inspectDeployedOffice,
} from './ppt-agent-release-preflight.mjs'

async function artifact(t) {
  const dist = await mkdtemp(resolve(tmpdir(), 'ppt-release-'))
  t.after(() => rm(dist, { recursive: true, force: true }))
  await mkdir(resolve(dist, 'assets'))
  await writeFile(resolve(dist, 'version.json'), '{"buildId":"release_123"}')
  await writeFile(
    resolve(dist, 'taskpane.html'),
    '<script src="/assets/taskpane-AbC_123.js"></script>',
  )
  await writeFile(resolve(dist, 'assets/taskpane-AbC_123.js'), 'const version="release_123"')
  await writeFile(
    resolve(dist, 'manifest.xml'),
    '<AppDomain>https://office.example</AppDomain><IconUrl DefaultValue="https://office.example/assets/icon.png"/><SourceLocation DefaultValue="https://office.example/taskpane.html"/>',
  )
  return dist
}

test('validates complete release artifact', async (t) => {
  const dist = await artifact(t)
  assert.deepEqual(await inspectOfficeBuild(dist, 'https://office.example'), {
    buildId: 'release_123',
    script: 'assets/taskpane-AbC_123.js',
    scriptSha256: createHash('sha256').update('const version="release_123"').digest('hex'),
  })
})

test('fails closed on mismatched metadata, origin and unhashed entry', async (t) => {
  const dist = await artifact(t)
  await writeFile(resolve(dist, 'version.json'), '{"buildId":"other"}')
  await assert.rejects(inspectOfficeBuild(dist, 'https://office.example'), /buildId differs/)
  await writeFile(resolve(dist, 'version.json'), '{"buildId":"release_123"}')
  await assert.rejects(
    inspectOfficeBuild(dist, 'https://other.example'),
    /manifest origin mismatch/,
  )
  await writeFile(resolve(dist, 'taskpane.html'), '<script src="/assets/taskpane.js"></script>')
  await assert.rejects(inspectOfficeBuild(dist, 'https://office.example'), /missing hashed/)
})

test('fails on source maps and unresolved connect policy', async (t) => {
  const dist = await artifact(t)
  await writeFile(
    resolve(dist, 'taskpane.html'),
    '__WISWORK_CONNECT_ORIGINS__<script src="/assets/taskpane-AbC_123.js"></script>',
  )
  await assert.rejects(inspectOfficeBuild(dist, 'https://office.example'), /unresolved/)
  await writeFile(
    resolve(dist, 'taskpane.html'),
    '<script src="/assets/taskpane-AbC_123.js"></script>',
  )
  await writeFile(resolve(dist, 'assets/taskpane-AbC_123.js.map'), '{}')
  await assert.rejects(inspectOfficeBuild(dist, 'https://office.example'), /source map/)
})

test('rejects conflicting Manifest URLs even when correct URLs are also present', async (t) => {
  const dist = await artifact(t)
  await writeFile(
    resolve(dist, 'manifest.xml'),
    '<AppDomain>https://office.example</AppDomain><SourceLocation DefaultValue="https://office.example/taskpane.html"/><SourceLocation DefaultValue="https://evil.example/taskpane.html"/>',
  )
  await assert.rejects(
    inspectOfficeBuild(dist, 'https://office.example'),
    /SourceLocation mismatch/,
  )
  await writeFile(
    resolve(dist, 'manifest.xml'),
    '<AppDomain>https://office.example</AppDomain><SourceLocation DefaultValue="https://office.example/taskpane.html"/><IconUrl DefaultValue="https://evil.example/assets/icon.png"/>',
  )
  await assert.rejects(
    inspectOfficeBuild(dist, 'https://office.example'),
    /IconUrl origin mismatch/,
  )
})

test('requires exact Relay health response and secure remote origin', async () => {
  const fetcher = async (url, options) => {
    assert.equal(url.href, 'https://relay.example/office-relay/health')
    assert.equal(options.cache, 'no-store')
    return new Response('ok')
  }
  await inspectRelayHealth('https://relay.example', fetcher)
  await assert.rejects(inspectRelayHealth('http://relay.example', fetcher), /requires HTTPS/)
  await assert.rejects(
    inspectRelayHealth('https://relay.example', async () => new Response('stale')),
    /health failed/,
  )
  await assert.rejects(
    inspectRelayHealth('https://relay.example', async () => new Response('ok', { status: 503 })),
    /health failed/,
  )
})

test('checks deployed version, HTML and immutable script as one build', async () => {
  const build = {
    buildId: 'release_123',
    script: 'assets/taskpane-AbC_123.js',
    scriptSha256: createHash('sha256').update('const version="release_123"').digest('hex'),
  }
  const assets = new Map([
    ['/version.json', '{"buildId":"release_123"}'],
    ['/taskpane.html', '<script src="/assets/taskpane-AbC_123.js"></script>'],
    ['/assets/taskpane-AbC_123.js', 'const version="release_123"'],
  ])
  const fetcher = async (url) =>
    new Response(assets.get(url.pathname) ?? '', { status: assets.has(url.pathname) ? 200 : 404 })
  await inspectDeployedOffice('https://office.example', build, fetcher)
  assets.set('/version.json', '{"buildId":"old"}')
  await assert.rejects(inspectDeployedOffice('https://office.example', build, fetcher), /differ/)
  assets.set('/version.json', '{"buildId":"release_123"}')
  assets.set('/assets/taskpane-AbC_123.js', 'const version="release_123";tampered=true')
  await assert.rejects(inspectDeployedOffice('https://office.example', build, fetcher), /differ/)
  assets.delete('/assets/taskpane-AbC_123.js')
  await assert.rejects(
    inspectDeployedOffice('https://office.example', build, fetcher),
    /unavailable/,
  )
})
