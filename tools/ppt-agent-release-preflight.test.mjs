import assert from 'node:assert/strict'
import { createHash } from 'node:crypto'
import { mkdtemp, mkdir, writeFile, rm } from 'node:fs/promises'
import { tmpdir } from 'node:os'
import { resolve } from 'node:path'
import { test } from 'node:test'
import { createServer } from 'node:http'
import { WebSocketServer } from 'ws'
import {
  inspectOfficeBuild,
  inspectRelayHealth,
  inspectDeployedOffice,
  inspectRelayPairing,
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
    htmlSha256: createHash('sha256')
      .update('<script src="/assets/taskpane-AbC_123.js"></script>')
      .digest('hex'),
  })
})

test('v2 pairing smoke negotiates and approves one presentation capability', async (t) => {
  const server = createServer()
  const ws = new WebSocketServer({ server, path: '/office-relay' })
  t.after(() => {
    ws.close()
    server.close()
  })
  server.listen(0, '127.0.0.1')
  await new Promise((resolveReady) => server.once('listening', resolveReady))
  let office
  let pc
  const seen = []
  ws.on('connection', (socket, request) => {
    if (request.headers.origin) {
      assert.equal(request.headers.origin, 'https://office.8-216-134-194.sslip.io')
      office = socket
    } else {
      assert.equal(request.headers.authorization, 'Bearer test-secret')
      pc = socket
    }
    socket.on('message', (bytes) => {
      const frame = JSON.parse(bytes.toString())
      seen.push(frame.type)
      assert.deepEqual(frame.capabilities, ['presentation.v1'])
      if (frame.type === 'office.create')
        socket.send(
          JSON.stringify({
            version: 2,
            type: 'office.created',
            pairing_id: 'pair',
            verification_code: '123456',
          }),
        )
      if (frame.type === 'pc.negotiate')
        socket.send(
          JSON.stringify({
            version: 2,
            type: 'pc.negotiated',
            pairing_version: 2,
            capabilities: ['presentation.v1'],
          }),
        )
      if (frame.type === 'pc.claim')
        socket.send(
          JSON.stringify({
            version: 2,
            type: 'pc.claimed',
            pairing_id: 'pair',
            capabilities: ['presentation.v1'],
          }),
        )
      if (frame.type === 'pc.approve') {
        socket.send(
          JSON.stringify({
            version: 2,
            type: 'pc.approved',
            session_id: 'session',
            capabilities: ['presentation.v1'],
          }),
        )
        office.send(
          JSON.stringify({
            version: 2,
            type: 'office.approved',
            session_id: 'session',
            capabilities: ['presentation.v1'],
          }),
        )
      }
    })
  })
  await inspectRelayPairing(`http://127.0.0.1:${server.address().port}`, 'test-secret')
  assert.ok(office)
  assert.ok(pc)
  assert.deepEqual(seen, ['office.create', 'pc.negotiate', 'pc.claim', 'pc.approve'])
})

test('pairing smoke fails closed on invalid destination and missing credential', async () => {
  await assert.rejects(inspectRelayPairing('http://relay.example', 'secret'), /requires HTTPS/)
  await assert.rejects(inspectRelayPairing('https://relay.example', ''), /requires PC token/)
})

test('pairing smoke rejects capability mismatch without disclosing credentials', async (t) => {
  const server = createServer()
  const ws = new WebSocketServer({ server, path: '/office-relay' })
  t.after(() => {
    ws.close()
    server.close()
  })
  server.listen(0, '127.0.0.1')
  await new Promise((resolveReady) => server.once('listening', resolveReady))
  ws.on('connection', (socket) =>
    socket.on('message', (bytes) => {
      const frame = JSON.parse(bytes.toString())
      if (frame.type === 'office.create')
        socket.send(
          JSON.stringify({
            version: 2,
            type: 'office.created',
            pairing_id: 'pair',
            verification_code: '123456',
          }),
        )
      if (frame.type === 'pc.negotiate')
        socket.send(
          JSON.stringify({
            version: 2,
            type: 'pc.negotiated',
            pairing_version: 2,
            capabilities: ['agent.v1'],
          }),
        )
    }),
  )
  await assert.rejects(
    inspectRelayPairing(`http://127.0.0.1:${server.address().port}`, 'test-secret'),
    (error) =>
      error.message.includes('capability negotiation mismatch') &&
      !error.message.includes('test-secret'),
  )
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
    htmlSha256: createHash('sha256')
      .update('<script src="/assets/taskpane-AbC_123.js"></script>')
      .digest('hex'),
  }
  const assets = new Map([
    ['/version.json', '{"buildId":"release_123"}'],
    ['/taskpane.html', '<script src="/assets/taskpane-AbC_123.js"></script>'],
    ['/assets/taskpane-AbC_123.js', 'const version="release_123"'],
  ])
  const fetcher = async (url) =>
    new Response(assets.get(url.pathname) ?? '', { status: assets.has(url.pathname) ? 200 : 404 })
  await inspectDeployedOffice('https://office.example', build, fetcher)
  assets.set(
    '/taskpane.html',
    '<script src="/assets/taskpane-AbC_123.js"></script><script src="https://evil.example/extra.js"></script>',
  )
  await assert.rejects(inspectDeployedOffice('https://office.example', build, fetcher), /differ/)
  assets.set('/taskpane.html', '<script src="/assets/taskpane-AbC_123.js"></script>')
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
