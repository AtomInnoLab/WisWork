import { randomUUID } from 'node:crypto'
import { fileURLToPath } from 'node:url'
import WebSocket from 'ws'

const OFFICE_ORIGIN = 'https://office.8-216-134-194.sslip.io'
const MAX_RESPONSE_BYTES = 16 * 1024 * 1024

function relayUrl(origin) {
  const url = new URL(origin)
  if (
    url.origin !== origin ||
    (url.protocol !== 'https:' &&
      !(url.protocol === 'http:' && ['localhost', '127.0.0.1', '[::1]'].includes(url.hostname)))
  )
    throw new Error('relay origin must be HTTPS or loopback HTTP')
  url.protocol = url.protocol === 'https:' ? 'wss:' : 'ws:'
  url.pathname = '/office-relay'
  return url
}

function mailbox(socket) {
  const frames = []
  const waiters = []
  let failure
  function deliver(error, frame) {
    const waiter = waiters.shift()
    if (waiter) error ? waiter.reject(error) : waiter.resolve(frame)
    else if (error) failure = error
    else frames.push(frame)
  }
  socket.on('message', (data) => {
    try {
      if (data.length > MAX_RESPONSE_BYTES) throw new Error('oversized relay frame')
      deliver(null, JSON.parse(data.toString()))
    } catch {
      deliver(new Error('invalid relay frame'))
    }
  })
  socket.on('close', () => deliver(new Error('relay connection closed')))
  socket.on('error', () => deliver(new Error('relay connection failed')))
  return async (timeoutMs) => {
    if (failure) throw failure
    const frame = frames.length
      ? frames.shift()
      : await new Promise((resolve, reject) => {
          const waiter = { resolve, reject }
          waiters.push(waiter)
          const timer = setTimeout(() => {
            const index = waiters.indexOf(waiter)
            if (index >= 0) waiters.splice(index, 1)
            reject(new Error('relay response timed out'))
          }, timeoutMs)
          waiter.resolve = (value) => {
            clearTimeout(timer)
            resolve(value)
          }
          waiter.reject = (error) => {
            clearTimeout(timer)
            reject(error)
          }
        })
    if (frame?.version !== 2 || typeof frame.type !== 'string')
      throw new Error('invalid relay protocol frame')
    if (frame.type === 'relay.error') throw new Error(`relay rejected request: ${frame.code}`)
    return frame
  }
}

async function expected(next, type, timeoutMs) {
  const frame = await next(timeoutMs)
  if (frame.type !== type) throw new Error(`expected ${type}, received ${frame.type}`)
  return frame
}

export function checkPcStatusResponse(response, projectId) {
  if (
    !response ||
    typeof response !== 'object' ||
    Array.isArray(response) ||
    response.projectId !== projectId ||
    !['planned', 'compiled', 'failed', 'pending'].includes(response.status) ||
    !Number.isSafeInteger(response.slideCount) ||
    response.slideCount < 0 ||
    !Array.isArray(response.slides) ||
    response.slides.length !== response.slideCount ||
    !Array.isArray(response.history)
  )
    throw new Error('PC presentation status response invalid')
  return { projectId: response.projectId, status: response.status, slideCount: response.slideCount }
}

export async function inspectPcBusiness(relayOrigin, documentId, projectId, options = {}) {
  if (
    typeof documentId !== 'string' ||
    !documentId.trim() ||
    documentId.length > 2048 ||
    typeof projectId !== 'string' ||
    !/^[A-Za-z0-9_-]{1,128}$/.test(projectId)
  )
    throw new Error('smoke requires a real document and project ID')
  const timeoutMs = options.timeoutMs ?? 120_000
  const socket = (options.connect ?? ((url, config) => new WebSocket(url, config)))(
    relayUrl(relayOrigin),
    { headers: { Origin: OFFICE_ORIGIN } },
  )
  const next = mailbox(socket)
  try {
    await new Promise((resolve, reject) => {
      const timer = setTimeout(() => reject(new Error('relay connection timed out')), 10_000)
      socket.once('open', () => {
        clearTimeout(timer)
        resolve()
      })
      socket.once('error', () => {
        clearTimeout(timer)
        reject(new Error('relay connection failed'))
      })
    })
    socket.send(
      JSON.stringify({
        version: 2,
        type: 'office.create',
        host: 'PowerPoint',
        capabilities: ['presentation.v1'],
      }),
    )
    const created = await expected(next, 'office.created', 10_000)
    if (!/^[0-9]{6}$/.test(created.verification_code))
      throw new Error('invalid relay pairing invitation')
    ;(
      options.onCode ??
      ((code) =>
        process.stdout.write(
          `Enter pairing code ${code} in the release-test WisWork PC and approve.\n`,
        ))
    )(created.verification_code)
    const approved = await expected(next, 'office.approved', timeoutMs)
    if (
      approved.session_id?.length < 1 ||
      approved.capability?.length < 1 ||
      !Array.isArray(approved.capabilities) ||
      !approved.capabilities.includes('presentation.v1')
    )
      throw new Error('PC did not negotiate presentation.v1')
    const requestId = randomUUID()
    socket.send(
      JSON.stringify({
        version: 2,
        type: 'office.request',
        session_id: approved.session_id,
        capability: approved.capability,
        request_id: requestId,
        capability_name: 'presentation.v1',
        body: { operation: 'status', documentId, projectId },
      }),
    )
    const start = await expected(next, 'relay.start', 30_000)
    if (
      start.request_id !== requestId ||
      start.session_id !== approved.session_id ||
      start.status !== 200 ||
      start.content_type !== 'application/json'
    )
      throw new Error('PC status request failed')
    const chunks = []
    let bytes = 0
    for (;;) {
      const frame = await next(30_000)
      if (frame.request_id !== requestId || frame.session_id !== approved.session_id)
        throw new Error('PC response identity mismatch')
      if (frame.type === 'relay.done') break
      if (
        frame.type !== 'relay.chunk' ||
        frame.sequence !== chunks.length ||
        typeof frame.data !== 'string' ||
        !/^(?:[A-Za-z0-9+/]{4})*(?:[A-Za-z0-9+/]{2}==|[A-Za-z0-9+/]{3}=)?$/.test(frame.data)
      )
        throw new Error('invalid PC response chunk')
      const chunk = Buffer.from(frame.data, 'base64')
      bytes += chunk.length
      if (bytes > MAX_RESPONSE_BYTES) throw new Error('PC response too large')
      chunks.push(chunk)
    }
    let response
    try {
      response = JSON.parse(Buffer.concat(chunks).toString('utf8'))
    } catch {
      throw new Error('invalid PC status JSON')
    }
    return checkPcStatusResponse(response, projectId)
  } finally {
    socket.terminate()
  }
}

if (process.argv[1] && fileURLToPath(import.meta.url) === process.argv[1]) {
  inspectPcBusiness(
    process.env.PPT_AGENT_SMOKE_RELAY_ORIGIN,
    process.env.PPT_AGENT_SMOKE_DOCUMENT_ID,
    process.env.PPT_AGENT_SMOKE_PROJECT_ID,
  )
    .then((result) =>
      process.stdout.write(
        `Real PC presentation status passed: ${result.status}, ${result.slideCount} slides.\n`,
      ),
    )
    .catch((error) => {
      process.stderr.write(`${error.message}\n`)
      process.exitCode = 1
    })
}
