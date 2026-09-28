import { lstat, readFile, readdir } from 'node:fs/promises'
import { createHash, randomUUID, randomBytes } from 'node:crypto'
import { resolve, dirname } from 'node:path'
import { fileURLToPath } from 'node:url'
import WebSocket from 'ws'

const defaultDist = resolve(dirname(fileURLToPath(import.meta.url)), '../apps/office-addin/dist')

function sha256(bytes) {
  return createHash('sha256').update(bytes).digest('hex')
}

function checkManifest(manifest, expectedOrigin) {
  const referenced = []
  const domains = [...manifest.matchAll(/<AppDomain>([^<]+)<\/AppDomain>/g)].map(
    (match) => match[1],
  )
  if (domains.length !== 1 || domains[0] !== expectedOrigin)
    throw new Error('manifest origin mismatch')
  for (const tag of ['SourceLocation', 'IconUrl', 'HighResolutionIconUrl', 'FunctionFile']) {
    const elements = [...manifest.matchAll(new RegExp(`<${tag}\\b[^>]*>`, 'g'))]
    if (tag === 'SourceLocation' && elements.length !== 1)
      throw new Error('manifest SourceLocation mismatch')
    for (const element of elements) {
      const value = element[0].match(/\bDefaultValue="([^"]+)"/)?.[1]
      if (!value) throw new Error(`manifest ${tag} missing URL`)
      let url
      try {
        url = new URL(value)
      } catch {
        throw new Error(`manifest ${tag} invalid URL`)
      }
      if (
        url.origin !== expectedOrigin ||
        (tag === 'SourceLocation' && url.pathname !== '/taskpane.html') ||
        (tag === 'IconUrl' && url.pathname !== '/assets/icon.png')
      )
        throw new Error(`manifest ${tag} origin mismatch`)
      if (url.hash) throw new Error(`manifest ${tag} invalid URL`)
      referenced.push(url.pathname.slice(1))
    }
  }
  return referenced
}

function referencedStylesheets(html) {
  const referenced = []
  for (const [element] of html.matchAll(/<link\b[^>]*>/g)) {
    const rel = element.match(/\brel=["']([^"']+)["']/)?.[1]
    if (!rel?.split(/\s+/).includes('stylesheet')) continue
    const href = element.match(/\bhref=["']([^"']+)["']/)?.[1]
    if (!href || !/^\/assets\/[A-Za-z0-9_.-]+\.css$/.test(href))
      throw new Error('invalid Office stylesheet reference')
    referenced.push(href.slice(1))
  }
  return referenced
}

function taskpaneScript(html) {
  const scripts = [...html.matchAll(/<script\b[^>]*>/g)].map(
    ([element]) => element.match(/\bsrc=["']([^"']+)["']/)?.[1],
  )
  const entries = scripts.filter((src) =>
    /^\/assets\/taskpane-[A-Za-z0-9_-]{7,}\.js$/.test(src || ''),
  )
  if (
    entries.length !== 1 ||
    scripts.some(
      (src) =>
        src !== entries[0] && src !== 'https://appsforoffice.microsoft.com/lib/1/hosted/office.js',
    ) ||
    scripts.filter((src) => src === 'https://appsforoffice.microsoft.com/lib/1/hosted/office.js')
      .length > 1
  )
    throw new Error('invalid Office script reference')
  return entries[0].slice(1)
}

function releaseAsset(path) {
  return (
    path === 'assets/icon.png' ||
    /^assets\/[A-Za-z0-9_.-]+-[A-Za-z0-9_-]{7,}\.(?:js|mjs|css|png|jpe?g|gif|svg|webp|woff2?)$/.test(
      path,
    )
  )
}

export async function inspectOfficeBuild(dist, expectedOrigin, expectedConfig) {
  const origin = new URL(expectedOrigin)
  if (origin.protocol !== 'https:' || origin.origin !== expectedOrigin)
    throw new Error('expected origin must be an HTTPS origin without a path')
  const entries = await readdir(dist, { recursive: true })
  const metadata = JSON.parse(await readFile(resolve(dist, 'version.json'), 'utf8'))
  if (!/^[A-Za-z0-9_.-]{3,96}$/.test(metadata.buildId || ''))
    throw new Error('invalid version.json buildId')
  for (const key of ['presentationRolloutPercent', 'diagnosticSamplePercent']) {
    if (!Number.isInteger(metadata[key]) || metadata[key] < 0 || metadata[key] > 100)
      throw new Error(`invalid version.json ${key}`)
    if (expectedConfig && metadata[key] !== expectedConfig[key])
      throw new Error(`release ${key} differs from expected configuration`)
  }
  const htmlBytes = await readFile(resolve(dist, 'taskpane.html'))
  const html = htmlBytes.toString('utf8')
  const manifest = await readFile(resolve(dist, 'manifest.xml'), 'utf8')
  const referenced = [...checkManifest(manifest, expectedOrigin), ...referencedStylesheets(html)]
  if (html.includes('__WISWORK_CONNECT_ORIGINS__')) throw new Error('unresolved connect policy')
  const entry = taskpaneScript(html)
  const script = await readFile(resolve(dist, entry))
  if (!script.toString('utf8').includes(metadata.buildId))
    throw new Error('buildId differs from compiled taskpane')
  const files = []
  let totalBytes = 0
  for (const path of entries.sort()) {
    const stat = await lstat(resolve(dist, path))
    if (stat.isDirectory()) continue
    if (path.endsWith('.map')) throw new Error('source map in release artifact')
    if (path.startsWith('assets/') && !releaseAsset(path))
      throw new Error(`unhashed Office runtime asset: ${path}`)
    if (
      !stat.isFile() ||
      (!['version.json', 'taskpane.html', 'manifest.xml'].includes(path) &&
        !/^assets\/[A-Za-z0-9_.-]+$/.test(path)) ||
      stat.size < 1 ||
      stat.size > 16 * 1024 * 1024
    )
      throw new Error(`invalid Office release file: ${path}`)
    totalBytes += stat.size
    if (files.length >= 128 || totalBytes > 64 * 1024 * 1024)
      throw new Error('Office release artifact too large')
    files.push({ path, size: stat.size, sha256: sha256(await readFile(resolve(dist, path))) })
  }
  const present = new Set(files.map((file) => file.path))
  for (const path of referenced)
    if (!present.has(path)) throw new Error(`missing referenced asset: ${path}`)
  return {
    buildId: metadata.buildId,
    presentationRolloutPercent: metadata.presentationRolloutPercent,
    diagnosticSamplePercent: metadata.diagnosticSamplePercent,
    script: entry,
    scriptSha256: sha256(script),
    htmlSha256: sha256(htmlBytes),
    files,
  }
}

export async function inspectRelayHealth(relayOrigin, fetcher = fetch) {
  const origin = new URL(relayOrigin)
  if (!['http:', 'https:'].includes(origin.protocol) || origin.origin !== relayOrigin)
    throw new Error('relay origin must be an HTTP(S) origin without a path')
  if (
    origin.protocol !== 'https:' &&
    !['localhost', '127.0.0.1', '[::1]'].includes(origin.hostname)
  )
    throw new Error('remote relay requires HTTPS')
  const response = await fetcher(new URL('/office-relay/health', origin), {
    cache: 'no-store',
    signal: AbortSignal.timeout(5_000),
  })
  if (!response.ok || (await response.text()).trim() !== 'ok')
    throw new Error(`relay health failed: ${response.status}`)
}

export async function inspectDeployedOffice(origin, build, fetcher = fetch) {
  if (!Array.isArray(build.files) || build.files.length < 4)
    throw new Error('incomplete Office release artifact')
  for (const file of build.files) {
    if (
      !file ||
      typeof file.path !== 'string' ||
      (!['version.json', 'taskpane.html', 'manifest.xml'].includes(file.path) &&
        !/^assets\/[A-Za-z0-9_.-]+$/.test(file.path)) ||
      !Number.isSafeInteger(file.size) ||
      file.size < 1 ||
      file.size > 16 * 1024 * 1024 ||
      !/^[a-f0-9]{64}$/.test(file.sha256)
    )
      throw new Error('invalid Office release file manifest')
    const url = new URL(`/${file.path}`, origin)
    const response = await fetcher(url, { cache: 'no-store', signal: AbortSignal.timeout(5_000) })
    if (!response.ok)
      throw new Error(`deployed asset unavailable: ${url.pathname} (${response.status})`)
    const cache = (response.headers.get('cache-control') || '')
      .toLowerCase()
      .split(',')
      .map((part) => part.trim())
    if (file.path.startsWith('assets/') && file.path !== 'assets/icon.png') {
      const maxAge = cache.find((part) => /^max-age=\d+$/.test(part))
      if (!cache.includes('immutable') || !maxAge || Number(maxAge.slice(8)) < 31_536_000)
        throw new Error(`deployed Office cache policy invalid: ${url.pathname}`)
    } else if (!cache.includes('no-store')) {
      throw new Error(`deployed Office cache policy invalid: ${url.pathname}`)
    }
    if (!response.body) throw new Error(`deployed asset unavailable: ${url.pathname}`)
    const reader = response.body.getReader()
    const chunks = []
    let bytes = 0
    try {
      while (true) {
        const { value, done } = await reader.read()
        if (done) break
        bytes += value.byteLength
        if (bytes > file.size)
          throw new Error('deployed Office assets differ from release artifact')
        chunks.push(value)
      }
    } finally {
      await reader.cancel().catch(() => {})
    }
    const body = Buffer.concat(chunks)
    if (bytes !== file.size || sha256(body) !== file.sha256)
      throw new Error('deployed Office assets differ from release artifact')
    if (file.path === 'version.json') {
      let metadata
      try {
        metadata = JSON.parse(body.toString('utf8'))
      } catch {
        throw new Error('invalid deployed version.json')
      }
      if (
        metadata.buildId !== build.buildId ||
        metadata.presentationRolloutPercent !== build.presentationRolloutPercent ||
        metadata.diagnosticSamplePercent !== build.diagnosticSamplePercent
      )
        throw new Error('deployed Office assets differ from release artifact')
    }
  }
}

const OFFICE_RELAY_ORIGIN = 'https://office.8-216-134-194.sslip.io'
const PAIRING_CAPABILITIES = [
  'presentation.v1',
  'presentation-attachments.v1',
  'presentation-assets.v1',
  'presentation-remote-images.v1',
  'presentation-webpages.v1',
  'presentation-asset-rights.v1',
]

function relaySocket(origin, options) {
  const url = new URL('/office-relay', origin)
  url.protocol = url.protocol === 'https:' ? 'wss:' : 'ws:'
  return new WebSocket(url, options)
}

function waitForOpen(socket, role, timeoutMs) {
  return new Promise((resolveOpen, reject) => {
    const timer = setTimeout(() => finish(new Error(`${role} relay upgrade timed out`)), timeoutMs)
    function finish(error) {
      clearTimeout(timer)
      socket.off('open', onOpen)
      socket.off('error', onError)
      socket.off('unexpected-response', onUpgradeError)
      if (error) reject(error)
      else resolveOpen()
    }
    function onOpen() {
      finish()
    }
    function onError() {
      finish(new Error(`${role} relay connection failed`))
    }
    function onUpgradeError() {
      finish(new Error(`${role} relay upgrade failed`))
    }
    socket.once('open', onOpen)
    socket.once('error', onError)
    socket.once('unexpected-response', onUpgradeError)
  })
}

function waitForFrame(socket, expected, timeoutMs) {
  return new Promise((resolveFrame, reject) => {
    const timer = setTimeout(
      () => finish(new Error(`relay pairing timed out at ${expected}`)),
      timeoutMs,
    )
    function finish(error, frame) {
      clearTimeout(timer)
      socket.off('message', onMessage)
      socket.off('close', onClose)
      socket.off('error', onError)
      if (error) reject(error)
      else resolveFrame(frame)
    }
    function onClose() {
      finish(new Error(`relay connection closed at ${expected}`))
    }
    function onError() {
      finish(new Error(`relay connection failed at ${expected}`))
    }
    function onMessage(data) {
      let frame
      try {
        frame = JSON.parse(data.toString())
      } catch {
        finish(new Error('invalid relay frame'))
        return
      }
      if (frame.type === 'relay.error') {
        finish(new Error(`relay pairing failed: ${frame.code}`))
        return
      }
      if (frame.type !== expected || frame.version !== 2) {
        finish(new Error(`unexpected relay frame at ${expected}`))
        return
      }
      finish(null, frame)
    }
    socket.on('message', onMessage)
    socket.once('close', onClose)
    socket.once('error', onError)
  })
}

function sendAndReceive(socket, frame, expected, timeoutMs) {
  const received = waitForFrame(socket, expected, timeoutMs)
  socket.send(JSON.stringify(frame))
  return received
}

function requireCapabilities(frame) {
  if (
    !Array.isArray(frame.capabilities) ||
    frame.capabilities.length !== PAIRING_CAPABILITIES.length ||
    frame.capabilities.some((capability, index) => capability !== PAIRING_CAPABILITIES[index])
  )
    throw new Error('relay capability negotiation mismatch')
}

function opaque(value) {
  return (
    typeof value === 'string' &&
    value.length > 0 &&
    value.length <= 512 &&
    /^[A-Za-z0-9_-]+$/.test(value)
  )
}

export async function inspectRelayPairing(relayOrigin, pcToken, options = {}) {
  const origin = new URL(relayOrigin)
  if (
    origin.origin !== relayOrigin ||
    (origin.protocol !== 'https:' &&
      !(
        origin.protocol === 'http:' && ['localhost', '127.0.0.1', '[::1]'].includes(origin.hostname)
      ))
  )
    throw new Error('relay pairing requires HTTPS or a loopback HTTP origin')
  if (typeof pcToken !== 'string' || !pcToken.trim())
    throw new Error('relay pairing requires PC token')
  const timeoutMs = options.timeoutMs ?? 5_000
  const connect = options.connect ?? relaySocket
  const office = connect(relayOrigin, { headers: { Origin: OFFICE_RELAY_ORIGIN } })
  let pc
  try {
    await waitForOpen(office, 'Office', timeoutMs)
    const created = await sendAndReceive(
      office,
      {
        version: 2,
        type: 'office.create',
        host: 'PowerPoint',
        capabilities: PAIRING_CAPABILITIES,
      },
      'office.created',
      timeoutMs,
    )
    if (!/^[0-9]{6}$/.test(created.verification_code) || typeof created.pairing_id !== 'string')
      throw new Error('invalid relay pairing invitation')
    pc = connect(relayOrigin, { headers: { Authorization: `Bearer ${pcToken}` } })
    await waitForOpen(pc, 'PC', timeoutMs)
    const negotiated = await sendAndReceive(
      pc,
      {
        version: 2,
        type: 'pc.negotiate',
        verification_code: created.verification_code,
        capabilities: PAIRING_CAPABILITIES,
      },
      'pc.negotiated',
      timeoutMs,
    )
    if (negotiated.pairing_version !== 2) throw new Error('relay pairing version mismatch')
    requireCapabilities(negotiated)
    const claimed = await sendAndReceive(
      pc,
      {
        version: 2,
        type: 'pc.claim',
        verification_code: created.verification_code,
        capabilities: PAIRING_CAPABILITIES,
      },
      'pc.claimed',
      timeoutMs,
    )
    if (claimed.pairing_id !== created.pairing_id) throw new Error('relay pairing ID mismatch')
    requireCapabilities(claimed)
    const officeApproved = waitForFrame(office, 'office.approved', timeoutMs)
    officeApproved.catch(() => {})
    const pcApproved = await sendAndReceive(
      pc,
      {
        version: 2,
        type: 'pc.approve',
        pairing_id: created.pairing_id,
        capabilities: PAIRING_CAPABILITIES,
      },
      'pc.approved',
      timeoutMs,
    )
    const approved = await officeApproved
    requireCapabilities(approved)
    requireCapabilities(pcApproved)
    if (!opaque(approved.session_id) || approved.session_id !== pcApproved.session_id)
      throw new Error('relay session mismatch')
    if (
      !opaque(approved.capability) ||
      !opaque(pcApproved.capability) ||
      approved.capability === pcApproved.capability
    )
      throw new Error('relay session capability credentials invalid')

    // Use the two smoke sockets to verify every presentation capability route without invoking PC operations.
    for (const capabilityName of PAIRING_CAPABILITIES) {
      const requestId = randomUUID()
      const challenge = randomBytes(16).toString('hex')
      const body = { operation: 'release_preflight', challenge }
      const forwardedFrame = waitForFrame(pc, 'relay.request', timeoutMs)
      office.send(
        JSON.stringify({
          version: 2,
          type: 'office.request',
          session_id: approved.session_id,
          capability: approved.capability,
          request_id: requestId,
          capability_name: capabilityName,
          body,
        }),
      )
      const forwarded = await forwardedFrame
      if (
        forwarded.session_id !== approved.session_id ||
        forwarded.request_id !== requestId ||
        forwarded.capability_name !== capabilityName ||
        !forwarded.body ||
        typeof forwarded.body !== 'object' ||
        Array.isArray(forwarded.body) ||
        Object.keys(forwarded.body).length !== 2 ||
        forwarded.body.operation !== body.operation ||
        forwarded.body.challenge !== body.challenge
      )
        throw new Error('relay request forwarding mismatch')

      const response = Buffer.from(challenge)
      const common = {
        version: 2,
        session_id: approved.session_id,
        capability: pcApproved.capability,
        request_id: requestId,
      }
      const startedFrame = waitForFrame(office, 'relay.start', timeoutMs)
      pc.send(
        JSON.stringify({
          ...common,
          type: 'pc.start',
          status: 200,
          content_type: 'application/octet-stream',
        }),
      )
      const started = await startedFrame
      if (
        started.session_id !== approved.session_id ||
        started.request_id !== requestId ||
        started.status !== 200 ||
        started.content_type !== 'application/octet-stream'
      )
        throw new Error('relay response start mismatch')
      const chunkFrame = waitForFrame(office, 'relay.chunk', timeoutMs)
      pc.send(
        JSON.stringify({
          ...common,
          type: 'pc.chunk',
          sequence: 0,
          data: response.toString('base64'),
        }),
      )
      const chunk = await chunkFrame
      if (
        chunk.session_id !== approved.session_id ||
        chunk.request_id !== requestId ||
        chunk.sequence !== 0 ||
        chunk.data !== response.toString('base64')
      )
        throw new Error('relay response chunk mismatch')
      const doneFrame = waitForFrame(office, 'relay.done', timeoutMs)
      pc.send(JSON.stringify({ ...common, type: 'pc.done' }))
      const done = await doneFrame
      if (done.session_id !== approved.session_id || done.request_id !== requestId)
        throw new Error('relay response completion mismatch')
    }
  } finally {
    office.terminate()
    pc?.terminate()
  }
}

async function main(args) {
  const options = Object.fromEntries(
    args.flatMap((arg, index) =>
      arg.startsWith('--') && args[index + 1] && !args[index + 1].startsWith('--')
        ? [[arg, args[index + 1]]]
        : [],
    ),
  )
  if (
    args.length !== Object.keys(options).length * 2 ||
    !options['--origin'] ||
    !options['--relay-origin'] ||
    Object.keys(options).some(
      (key) =>
        ![
          '--origin',
          '--relay-origin',
          '--dist',
          '--deployed',
          '--pairing',
          '--rollout-percent',
          '--diagnostic-sample-percent',
        ].includes(key),
    ) ||
    (options['--deployed'] && options['--deployed'] !== '1') ||
    (options['--pairing'] && options['--pairing'] !== '1') ||
    (options['--deployed'] &&
      (!options['--rollout-percent'] || !options['--diagnostic-sample-percent'])) ||
    Boolean(options['--rollout-percent']) !== Boolean(options['--diagnostic-sample-percent']) ||
    [options['--rollout-percent'], options['--diagnostic-sample-percent']].some(
      (value) => value !== undefined && !/^(?:0|[1-9]\d?|100)$/.test(value),
    )
  )
    throw new Error(
      'usage: node tools/ppt-agent-release-preflight.mjs --origin https://office.example --relay-origin https://relay.example [--dist path] [--deployed 1 --rollout-percent 0..100 --diagnostic-sample-percent 0..100] [--pairing 1]',
    )
  const build = await inspectOfficeBuild(
    resolve(options['--dist'] || defaultDist),
    options['--origin'],
    options['--rollout-percent'] !== undefined &&
      options['--diagnostic-sample-percent'] !== undefined
      ? {
          presentationRolloutPercent: Number(options['--rollout-percent']),
          diagnosticSamplePercent: Number(options['--diagnostic-sample-percent']),
        }
      : undefined,
  )
  if (options['--deployed']) await inspectDeployedOffice(options['--origin'], build)
  await inspectRelayHealth(options['--relay-origin'])
  if (options['--pairing']) {
    if (!options['--deployed']) throw new Error('pairing smoke requires --deployed 1')
    await inspectRelayPairing(options['--relay-origin'], process.env.PPT_AGENT_RELEASE_PC_TOKEN)
  }
  process.stdout.write(
    `PPT Agent release preflight passed: build ${build.buildId}, rollout ${build.presentationRolloutPercent}%, diagnostics ${build.diagnosticSamplePercent}%, ${build.files.length} artifact files${options['--deployed'] ? ' matched deployed bytes' : ' checked locally'}, Relay healthy${options['--pairing'] ? ', v2 presentation capability routes verified' : ''}\n`,
  )
}

if (process.argv[1] && resolve(process.argv[1]) === fileURLToPath(import.meta.url)) {
  main(process.argv.slice(2)).catch((error) => {
    process.stderr.write(`PPT Agent release preflight failed: ${error.message}\n`)
    process.exitCode = 1
  })
}
