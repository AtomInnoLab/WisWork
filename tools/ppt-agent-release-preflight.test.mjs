import assert from 'node:assert/strict'
import { createHash } from 'node:crypto'
import { mkdtemp, mkdir, writeFile, rm, symlink } from 'node:fs/promises'
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

const VERSION =
  '{"buildId":"release_123","presentationRolloutPercent":25,"diagnosticSamplePercent":10}'
const TEAM_START =
  '<script src="https://appsforoffice.microsoft.com/lib/1/hosted/office.js"></script><script src="/assets/teamAuthStart-AbC_123.js"></script><link rel="modulepreload" href="/assets/team-auth-config-AbC_123.js">'
const TEAM_CALLBACK =
  '<script src="https://appsforoffice.microsoft.com/lib/1/hosted/office.js"></script><script src="/assets/teamAuthCallback-AbC_123.js"></script><link rel="modulepreload" href="/assets/team-auth-config-AbC_123.js">'

const PRESENTATION_CAPABILITIES = [
  'presentation.v1',
  'presentation-attachments.v1',
  'presentation-assets.v1',
  'presentation-remote-images.v1',
  'presentation-webpages.v1',
  'presentation-asset-rights.v1',
]

async function artifact(t) {
  const dist = await mkdtemp(resolve(tmpdir(), 'ppt-release-'))
  t.after(() => rm(dist, { recursive: true, force: true }))
  await mkdir(resolve(dist, 'assets'))
  await writeFile(resolve(dist, 'version.json'), VERSION)
  await writeFile(
    resolve(dist, 'taskpane.html'),
    '<script src="/assets/taskpane-AbC_123.js"></script>',
  )
  await writeFile(resolve(dist, 'assets/taskpane-AbC_123.js'), 'const version="release_123"')
  await writeFile(resolve(dist, 'assets/taskpane-AbC_123.css'), 'body{color:#123456}')
  await writeFile(resolve(dist, 'assets/worker-AbC_123.js'), 'self.onmessage=()=>{}')
  await writeFile(resolve(dist, 'team-auth-start.html'), TEAM_START)
  await writeFile(resolve(dist, 'team-auth-callback.html'), TEAM_CALLBACK)
  await writeFile(resolve(dist, 'assets/teamAuthStart-AbC_123.js'), 'start()')
  await writeFile(resolve(dist, 'assets/teamAuthCallback-AbC_123.js'), 'callback()')
  await writeFile(resolve(dist, 'assets/team-auth-config-AbC_123.js'), 'config()')
  await writeFile(resolve(dist, 'assets/icon.png'), Buffer.from([137, 80, 78, 71]))
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
    presentationRolloutPercent: 25,
    diagnosticSamplePercent: 10,
    script: 'assets/taskpane-AbC_123.js',
    scriptSha256: createHash('sha256').update('const version="release_123"').digest('hex'),
    htmlSha256: createHash('sha256')
      .update('<script src="/assets/taskpane-AbC_123.js"></script>')
      .digest('hex'),
    files: [
      ['assets/icon.png', Buffer.from([137, 80, 78, 71])],
      ['assets/taskpane-AbC_123.css', Buffer.from('body{color:#123456}')],
      ['assets/taskpane-AbC_123.js', Buffer.from('const version="release_123"')],
      ['assets/team-auth-config-AbC_123.js', Buffer.from('config()')],
      ['assets/teamAuthCallback-AbC_123.js', Buffer.from('callback()')],
      ['assets/teamAuthStart-AbC_123.js', Buffer.from('start()')],
      ['assets/worker-AbC_123.js', Buffer.from('self.onmessage=()=>{}')],
      [
        'manifest.xml',
        Buffer.from(
          '<AppDomain>https://office.example</AppDomain><IconUrl DefaultValue="https://office.example/assets/icon.png"/><SourceLocation DefaultValue="https://office.example/taskpane.html"/>',
        ),
      ],
      ['taskpane.html', Buffer.from('<script src="/assets/taskpane-AbC_123.js"></script>')],
      ['team-auth-callback.html', Buffer.from(TEAM_CALLBACK)],
      ['team-auth-start.html', Buffer.from(TEAM_START)],
      ['version.json', Buffer.from(VERSION)],
    ].map(([path, bytes]) => ({
      path,
      size: bytes.length,
      sha256: createHash('sha256').update(bytes).digest('hex'),
    })),
  })
})

test('release configuration must match the operator expectation', async (t) => {
  const dist = await artifact(t)
  await inspectOfficeBuild(dist, 'https://office.example', {
    presentationRolloutPercent: 25,
    diagnosticSamplePercent: 10,
  })
  await assert.rejects(
    inspectOfficeBuild(dist, 'https://office.example', {
      presentationRolloutPercent: 100,
      diagnosticSamplePercent: 10,
    }),
    /presentationRolloutPercent differs/,
  )
  await assert.rejects(
    inspectOfficeBuild(dist, 'https://office.example', {
      presentationRolloutPercent: 25,
      diagnosticSamplePercent: 100,
    }),
    /diagnosticSamplePercent differs/,
  )
  await writeFile(resolve(dist, 'version.json'), '{"buildId":"release_123"}')
  await assert.rejects(inspectOfficeBuild(dist, 'https://office.example'), /invalid version.json/)
})

test('v2 pairing smoke negotiates and routes every presentation capability', async (t) => {
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
  let corruptChunk = false
  const seen = []
  const routed = []
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
      if (frame.capabilities) assert.deepEqual(frame.capabilities, PRESENTATION_CAPABILITIES)
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
            capabilities: PRESENTATION_CAPABILITIES,
          }),
        )
      if (frame.type === 'pc.claim')
        socket.send(
          JSON.stringify({
            version: 2,
            type: 'pc.claimed',
            pairing_id: 'pair',
            capabilities: PRESENTATION_CAPABILITIES,
          }),
        )
      if (frame.type === 'pc.approve') {
        socket.send(
          JSON.stringify({
            version: 2,
            type: 'pc.approved',
            session_id: 'session',
            capability: 'pcCredential',
            capabilities: PRESENTATION_CAPABILITIES,
          }),
        )
        office.send(
          JSON.stringify({
            version: 2,
            type: 'office.approved',
            session_id: 'session',
            capability: 'officeCredential',
            capabilities: PRESENTATION_CAPABILITIES,
          }),
        )
      }
      if (frame.type === 'office.request') {
        assert.equal(frame.capability, 'officeCredential')
        assert.ok(PRESENTATION_CAPABILITIES.includes(frame.capability_name))
        routed.push(frame.capability_name)
        assert.equal(frame.body.operation, 'release_preflight')
        pc.send(
          JSON.stringify({
            version: 2,
            type: 'relay.request',
            session_id: frame.session_id,
            request_id: frame.request_id,
            capability_name: frame.capability_name,
            body: frame.body,
          }),
        )
      }
      if (frame.type === 'pc.start' || frame.type === 'pc.chunk' || frame.type === 'pc.done') {
        assert.equal(frame.capability, 'pcCredential')
        office.send(
          JSON.stringify({
            version: 2,
            type: `relay.${frame.type.slice(3)}`,
            session_id: frame.session_id,
            request_id: frame.request_id,
            ...(frame.type === 'pc.start'
              ? { status: frame.status, content_type: frame.content_type }
              : {}),
            ...(frame.type === 'pc.chunk'
              ? { sequence: corruptChunk ? 1 : frame.sequence, data: frame.data }
              : {}),
          }),
        )
      }
    })
  })
  await inspectRelayPairing(`http://127.0.0.1:${server.address().port}`, 'test-secret')
  assert.ok(office)
  assert.ok(pc)
  assert.deepEqual(routed, PRESENTATION_CAPABILITIES)
  assert.deepEqual(seen, [
    'office.create',
    'pc.negotiate',
    'pc.claim',
    'pc.approve',
    ...PRESENTATION_CAPABILITIES.flatMap(() => [
      'office.request',
      'pc.start',
      'pc.chunk',
      'pc.done',
    ]),
  ])
  corruptChunk = true
  await assert.rejects(
    inspectRelayPairing(`http://127.0.0.1:${server.address().port}`, 'test-secret'),
    /relay response chunk mismatch/,
  )
})

test('pairing smoke fails closed on invalid destination and missing credential', async () => {
  await assert.rejects(inspectRelayPairing('http://relay.example', 'secret'), /requires HTTPS/)
  await assert.rejects(inspectRelayPairing('https://relay.example', ''), /requires PC token/)
})

test('pairing smoke requires distinct nonempty session credentials', async (t) => {
  const server = createServer()
  const ws = new WebSocketServer({ server, path: '/office-relay' })
  t.after(() => {
    ws.close()
    server.close()
  })
  server.listen(0, '127.0.0.1')
  await new Promise((resolveReady) => server.once('listening', resolveReady))
  let office
  let caseNumber = 0
  ws.on('connection', (socket, request) => {
    if (request.headers.origin) office = socket
    socket.on('message', (bytes) => {
      const { type } = JSON.parse(bytes.toString())
      if (type === 'office.create')
        socket.send(
          JSON.stringify({
            version: 2,
            type: 'office.created',
            pairing_id: 'pair',
            verification_code: '123456',
          }),
        )
      if (type === 'pc.negotiate')
        socket.send(
          JSON.stringify({
            version: 2,
            type: 'pc.negotiated',
            pairing_version: 2,
            capabilities: PRESENTATION_CAPABILITIES,
          }),
        )
      if (type === 'pc.claim')
        socket.send(
          JSON.stringify({
            version: 2,
            type: 'pc.claimed',
            pairing_id: 'pair',
            capabilities: PRESENTATION_CAPABILITIES,
          }),
        )
      if (type === 'pc.approve') {
        const capability = caseNumber === 0 ? undefined : 'sameCredential'
        socket.send(
          JSON.stringify({
            version: 2,
            type: 'pc.approved',
            session_id: 'session',
            capability,
            capabilities: PRESENTATION_CAPABILITIES,
          }),
        )
        office.send(
          JSON.stringify({
            version: 2,
            type: 'office.approved',
            session_id: 'session',
            capability: 'sameCredential',
            capabilities: PRESENTATION_CAPABILITIES,
          }),
        )
      }
    })
  })
  for (caseNumber = 0; caseNumber < 2; caseNumber++) {
    await assert.rejects(
      inspectRelayPairing(`http://127.0.0.1:${server.address().port}`, 'test-secret'),
      /capability credentials invalid/,
    )
  }
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
            capabilities: PRESENTATION_CAPABILITIES.slice(0, -1),
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
  await writeFile(resolve(dist, 'version.json'), VERSION.replace('release_123', 'other'))
  await assert.rejects(inspectOfficeBuild(dist, 'https://office.example'), /buildId differs/)
  await writeFile(resolve(dist, 'version.json'), VERSION)
  await assert.rejects(
    inspectOfficeBuild(dist, 'https://other.example'),
    /manifest origin mismatch/,
  )
  await writeFile(resolve(dist, 'taskpane.html'), '<script src="/assets/taskpane.js"></script>')
  await assert.rejects(
    inspectOfficeBuild(dist, 'https://office.example'),
    /invalid Office script reference/,
  )
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

test('requires hashed runtime assets and rejects extra remote HTML scripts', async (t) => {
  const dist = await artifact(t)
  await writeFile(resolve(dist, 'assets/runtime.js'), 'self.onmessage=()=>{}')
  await assert.rejects(
    inspectOfficeBuild(dist, 'https://office.example'),
    /unhashed Office runtime asset/,
  )
  await rm(resolve(dist, 'assets/runtime.js'))
  await writeFile(resolve(dist, 'assets/theme.css'), 'body{color:red}')
  await assert.rejects(
    inspectOfficeBuild(dist, 'https://office.example'),
    /unhashed Office runtime asset/,
  )
  await rm(resolve(dist, 'assets/theme.css'))
  await writeFile(
    resolve(dist, 'taskpane.html'),
    '<script src="/assets/taskpane-AbC_123.js"></script><script src="https://other.example/extra.js"></script>',
  )
  await assert.rejects(
    inspectOfficeBuild(dist, 'https://office.example'),
    /invalid Office script reference/,
  )
})

test('checks both team authentication pages and their hashed dependencies', async (t) => {
  const dist = await artifact(t)
  await inspectOfficeBuild(dist, 'https://office.example')
  await writeFile(
    resolve(dist, 'team-auth-start.html'),
    TEAM_START.replace('/assets/teamAuthStart-AbC_123.js', 'https://evil.example/start.js'),
  )
  await assert.rejects(
    inspectOfficeBuild(dist, 'https://office.example'),
    /invalid Office team-auth-start.html script reference/,
  )
  await writeFile(resolve(dist, 'team-auth-start.html'), TEAM_START)
  await rm(resolve(dist, 'assets/team-auth-config-AbC_123.js'))
  await assert.rejects(
    inspectOfficeBuild(dist, 'https://office.example'),
    /missing referenced asset/,
  )
  await rm(resolve(dist, 'team-auth-callback.html'))
  await assert.rejects(inspectOfficeBuild(dist, 'https://office.example'), /ENOENT/)
})

test('checks taskpane module preloads as release dependencies', async (t) => {
  const dist = await artifact(t)
  const taskpane =
    '<script src="/assets/taskpane-AbC_123.js"></script><link rel="modulepreload" href="/assets/shared-AbC_123.js">'
  await writeFile(resolve(dist, 'taskpane.html'), taskpane)
  await assert.rejects(
    inspectOfficeBuild(dist, 'https://office.example'),
    /missing referenced asset: assets\/shared-AbC_123.js/,
  )
  await writeFile(resolve(dist, 'assets/shared-AbC_123.js'), 'export const shared = true')
  await inspectOfficeBuild(dist, 'https://office.example')
  await writeFile(
    resolve(dist, 'taskpane.html'),
    taskpane.replace('/assets/shared-AbC_123.js', 'https://other.example/shared.js'),
  )
  await assert.rejects(
    inspectOfficeBuild(dist, 'https://office.example'),
    /invalid Office taskpane.html preload reference/,
  )
})

test('rejects unlisted and symlinked release files', async (t) => {
  const dist = await artifact(t)
  await writeFile(resolve(dist, 'assets/unexpected.map'), '{}')
  await assert.rejects(inspectOfficeBuild(dist, 'https://office.example'), /source map/)
  await rm(resolve(dist, 'assets/unexpected.map'))
  await symlink('../version.json', resolve(dist, 'assets/linked-AbC_123.js'))
  await assert.rejects(
    inspectOfficeBuild(dist, 'https://office.example'),
    /invalid Office release file/,
  )
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

test('rejects HTML stylesheet and Manifest resources absent from the release', async (t) => {
  const dist = await artifact(t)
  await writeFile(
    resolve(dist, 'taskpane.html'),
    '<script src="/assets/taskpane-AbC_123.js"></script><link rel="stylesheet" href="/assets/missing.css">',
  )
  await assert.rejects(
    inspectOfficeBuild(dist, 'https://office.example'),
    /missing referenced asset/,
  )
  await writeFile(
    resolve(dist, 'taskpane.html'),
    '<script src="/assets/taskpane-AbC_123.js"></script><link href="/assets/taskpane-AbC_123.css" rel="stylesheet">',
  )
  await inspectOfficeBuild(dist, 'https://office.example')
  await writeFile(
    resolve(dist, 'manifest.xml'),
    '<AppDomain>https://office.example</AppDomain><IconUrl DefaultValue="https://office.example/assets/icon.png"/><HighResolutionIconUrl DefaultValue="https://office.example/assets/missing.png"/><SourceLocation DefaultValue="https://office.example/taskpane.html"/>',
  )
  await assert.rejects(
    inspectOfficeBuild(dist, 'https://office.example'),
    /missing referenced asset/,
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
  const assets = new Map([
    ['/version.json', VERSION],
    ['/taskpane.html', '<script src="/assets/taskpane-AbC_123.js"></script>'],
    ['/team-auth-start.html', TEAM_START],
    ['/team-auth-callback.html', TEAM_CALLBACK],
    ['/assets/taskpane-AbC_123.js', 'const version="release_123"'],
    ['/assets/teamAuthStart-AbC_123.js', 'start()'],
    ['/assets/teamAuthCallback-AbC_123.js', 'callback()'],
    ['/assets/team-auth-config-AbC_123.js', 'config()'],
    ['/assets/taskpane-AbC_123.css', 'body{color:#123456}'],
    ['/assets/worker-AbC_123.js', 'self.onmessage=()=>{}'],
    ['/assets/icon.png', Buffer.from([137, 80, 78, 71])],
    ['/manifest.xml', '<AppDomain>https://office.example</AppDomain>'],
  ])
  const build = {
    buildId: 'release_123',
    presentationRolloutPercent: 25,
    diagnosticSamplePercent: 10,
    script: 'assets/taskpane-AbC_123.js',
    scriptSha256: createHash('sha256').update('const version="release_123"').digest('hex'),
    htmlSha256: createHash('sha256')
      .update('<script src="/assets/taskpane-AbC_123.js"></script>')
      .digest('hex'),
    files: [...assets].map(([path, value]) => ({
      path: path.slice(1),
      size: Buffer.byteLength(value),
      sha256: createHash('sha256').update(value).digest('hex'),
    })),
  }
  const cacheHeaders = (path) => ({
    'Cache-Control':
      path.startsWith('/assets/') && path !== '/assets/icon.png'
        ? 'public, max-age=31536000, immutable'
        : 'no-store',
  })
  let omitCacheFor
  const fetcher = async (url) =>
    new Response(assets.get(url.pathname) ?? '', {
      status: assets.has(url.pathname) ? 200 : 404,
      headers: url.pathname === omitCacheFor ? {} : cacheHeaders(url.pathname),
    })
  await inspectDeployedOffice('https://office.example', build, fetcher)
  assets.set(
    '/team-auth-callback.html',
    TEAM_CALLBACK.replace('teamAuthCallback-AbC_123.js', 'teamAuthCallback-XbC_123.js'),
  )
  await assert.rejects(inspectDeployedOffice('https://office.example', build, fetcher), /differ/)
  assets.set('/team-auth-callback.html', TEAM_CALLBACK)
  omitCacheFor = '/version.json'
  await assert.rejects(
    inspectDeployedOffice('https://office.example', build, fetcher),
    /cache policy/,
  )
  omitCacheFor = '/assets/taskpane-AbC_123.js'
  await assert.rejects(
    inspectDeployedOffice('https://office.example', build, fetcher),
    /cache policy/,
  )
  omitCacheFor = undefined
  assets.set(
    '/taskpane.html',
    '<script src="/assets/taskpane-AbC_123.js"></script><script src="https://evil.example/extra.js"></script>',
  )
  await assert.rejects(inspectDeployedOffice('https://office.example', build, fetcher), /differ/)
  assets.set('/taskpane.html', '<script src="/assets/taskpane-AbC_123.js"></script>')
  assets.set('/version.json', '{"buildId":"old"}')
  await assert.rejects(inspectDeployedOffice('https://office.example', build, fetcher), /differ/)
  assets.set('/version.json', VERSION)
  assets.set('/assets/taskpane-AbC_123.js', 'const version="release_123";tampered=true')
  await assert.rejects(inspectDeployedOffice('https://office.example', build, fetcher), /differ/)
  assets.delete('/assets/taskpane-AbC_123.js')
  await assert.rejects(
    inspectDeployedOffice('https://office.example', build, fetcher),
    /unavailable/,
  )
  assets.set('/assets/taskpane-AbC_123.js', 'const version="release_123"')
  assets.delete('/assets/taskpane-AbC_123.css')
  await assert.rejects(
    inspectDeployedOffice('https://office.example', build, fetcher),
    /unavailable/,
  )
  assets.set('/assets/taskpane-AbC_123.css', 'body{color:#123456}')
  assets.set('/assets/worker-AbC_123.js', 'self.onmessage=()=>{throw Error("changed")}')
  await assert.rejects(inspectDeployedOffice('https://office.example', build, fetcher), /differ/)
  assets.set('/assets/worker-AbC_123.js', 'x'.repeat(2 * 1024 * 1024))
  await assert.rejects(inspectDeployedOffice('https://office.example', build, fetcher), /differ/)
})
