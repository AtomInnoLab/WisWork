import { readFile, readdir } from 'node:fs/promises'
import { createHash } from 'node:crypto'
import { resolve, dirname } from 'node:path'
import { fileURLToPath } from 'node:url'
import WebSocket from 'ws'

const defaultDist = resolve(dirname(fileURLToPath(import.meta.url)), '../apps/office-addin/dist')

function sha256(bytes) {
  return createHash('sha256').update(bytes).digest('hex')
}

function checkManifest(manifest, expectedOrigin) {
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
    }
  }
}

export async function inspectOfficeBuild(dist, expectedOrigin) {
  const origin = new URL(expectedOrigin)
  if (origin.protocol !== 'https:' || origin.origin !== expectedOrigin)
    throw new Error('expected origin must be an HTTPS origin without a path')
  const files = await readdir(dist, { recursive: true })
  const metadata = JSON.parse(await readFile(resolve(dist, 'version.json'), 'utf8'))
  if (!/^[A-Za-z0-9_.-]{3,96}$/.test(metadata.buildId || ''))
    throw new Error('invalid version.json buildId')
  const htmlBytes = await readFile(resolve(dist, 'taskpane.html'))
  const html = htmlBytes.toString('utf8')
  const manifest = await readFile(resolve(dist, 'manifest.xml'), 'utf8')
  checkManifest(manifest, expectedOrigin)
  if (html.includes('__WISWORK_CONNECT_ORIGINS__')) throw new Error('unresolved connect policy')
  const scripts = [...html.matchAll(/<script\b[^>]*\bsrc="(\/assets\/[^"?]+\.js)"/g)].map((match) =>
    match[1].slice(1),
  )
  if (scripts.length !== 1 || !/^assets\/taskpane-[A-Za-z0-9_-]+\.js$/.test(scripts[0]))
    throw new Error('missing hashed taskpane entry')
  const script = await readFile(resolve(dist, scripts[0]))
  if (!script.toString('utf8').includes(metadata.buildId))
    throw new Error('buildId differs from compiled taskpane')
  for (const path of files) {
    if (path.endsWith('.map')) throw new Error('source map in release artifact')
  }
  return {
    buildId: metadata.buildId,
    script: scripts[0],
    scriptSha256: sha256(script),
    htmlSha256: sha256(htmlBytes),
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
  const requested = [
    new URL('/version.json', origin),
    new URL('/taskpane.html', origin),
    new URL(`/${build.script}`, origin),
  ]
  const responses = []
  for (const url of requested) {
    const response = await fetcher(url, { cache: 'no-store', signal: AbortSignal.timeout(5_000) })
    if (!response.ok)
      throw new Error(`deployed asset unavailable: ${url.pathname} (${response.status})`)
    responses.push(Buffer.from(await response.arrayBuffer()))
  }
  let metadata
  try {
    metadata = JSON.parse(responses[0].toString('utf8'))
  } catch {
    throw new Error('invalid deployed version.json')
  }
  if (
    metadata.buildId !== build.buildId ||
    sha256(responses[1]) !== build.htmlSha256 ||
    sha256(responses[2]) !== build.scriptSha256
  )
    throw new Error('deployed Office assets differ from release artifact')
}

const OFFICE_RELAY_ORIGIN = 'https://office.8-216-134-194.sslip.io'
const PAIRING_CAPABILITY = 'presentation.v1'

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
    frame.capabilities.length !== 1 ||
    frame.capabilities[0] !== PAIRING_CAPABILITY
  )
    throw new Error('relay capability negotiation mismatch')
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
        capabilities: [PAIRING_CAPABILITY],
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
        capabilities: [PAIRING_CAPABILITY],
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
        capabilities: [PAIRING_CAPABILITY],
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
        capabilities: [PAIRING_CAPABILITY],
      },
      'pc.approved',
      timeoutMs,
    )
    const approved = await officeApproved
    requireCapabilities(approved)
    requireCapabilities(pcApproved)
    if (!approved.session_id || approved.session_id !== pcApproved.session_id)
      throw new Error('relay session mismatch')
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
      (key) => !['--origin', '--relay-origin', '--dist', '--deployed', '--pairing'].includes(key),
    ) ||
    (options['--deployed'] && options['--deployed'] !== '1') ||
    (options['--pairing'] && options['--pairing'] !== '1')
  )
    throw new Error(
      'usage: node tools/ppt-agent-release-preflight.mjs --origin https://office.example --relay-origin https://relay.example [--dist path] [--deployed 1] [--pairing 1]',
    )
  const build = await inspectOfficeBuild(
    resolve(options['--dist'] || defaultDist),
    options['--origin'],
  )
  if (options['--deployed']) await inspectDeployedOffice(options['--origin'], build)
  await inspectRelayHealth(options['--relay-origin'])
  if (options['--pairing']) {
    if (!options['--deployed']) throw new Error('pairing smoke requires --deployed 1')
    await inspectRelayPairing(options['--relay-origin'], process.env.PPT_AGENT_RELEASE_PC_TOKEN)
  }
  process.stdout.write(
    `PPT Agent release preflight passed: build ${build.buildId}, ${build.script}, Relay healthy${options['--pairing'] ? ', v2 pairing verified' : ''}\n`,
  )
}

if (process.argv[1] && resolve(process.argv[1]) === fileURLToPath(import.meta.url)) {
  main(process.argv.slice(2)).catch((error) => {
    process.stderr.write(`PPT Agent release preflight failed: ${error.message}\n`)
    process.exitCode = 1
  })
}
