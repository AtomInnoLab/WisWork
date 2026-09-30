import { createHash, randomBytes, randomUUID } from 'node:crypto'
import { fileURLToPath } from 'node:url'
import pngjs from 'pngjs'
import WebSocket from 'ws'

const { PNG } = pngjs

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

function checkAttachmentList(response) {
  if (
    !response ||
    !Array.isArray(response.attachments) ||
    response.attachments.length > 32 ||
    (response.nextAfter !== undefined &&
      (response.attachments.length !== 32 ||
        response.nextAfter !== response.attachments.at(-1)?.attachmentId)) ||
    response.attachments.some(
      (item, index) =>
        !/^[a-f0-9]{64}$/.test(item?.attachmentId) ||
        (index > 0 && item.attachmentId <= response.attachments[index - 1].attachmentId),
    )
  )
    throw new Error('PC attachment listing invalid')
  return response.attachments.length
}

function checkTextAttachment(metadata, read, id) {
  if (
    metadata?.attachmentId !== id ||
    metadata.status !== 'ready' ||
    metadata.kind !== 'text' ||
    read?.attachmentId !== id ||
    read.offset !== 0 ||
    read.sourceUri !== `attachment:${id}` ||
    !Number.isSafeInteger(read.totalChars) ||
    read.totalChars < 1 ||
    read.totalChars !== metadata.totalChars ||
    typeof read.text !== 'string' ||
    !read.text.trim() ||
    read.text.length > 8000 ||
    read.text.length !== Math.min(read.totalChars, 8000)
  )
    throw new Error('PC text attachment read invalid')
}

function checkImageAttachment(metadata, asset, id) {
  if (
    metadata?.attachmentId !== id ||
    metadata.status !== 'ready' ||
    metadata.kind !== 'image' ||
    asset?.id !== id ||
    asset.mime !== 'image/png' ||
    !Number.isSafeInteger(asset.width) ||
    asset.width < 1 ||
    asset.width !== metadata.width ||
    !Number.isSafeInteger(asset.height) ||
    asset.height < 1 ||
    asset.height !== metadata.height ||
    typeof asset.base64 !== 'string' ||
    !/^(?:[A-Za-z0-9+/]{4})*(?:[A-Za-z0-9+/]{2}==|[A-Za-z0-9+/]{3}=)?$/.test(asset.base64)
  )
    throw new Error('PC image attachment read invalid')
  const bytes = Buffer.from(asset.base64, 'base64')
  if (
    !bytes.length ||
    bytes.length > 10 * 1024 * 1024 ||
    bytes.toString('base64') !== asset.base64 ||
    createHash('sha256').update(bytes).digest('hex') !== metadata.assetSha256
  )
    throw new Error('PC image attachment digest mismatch')
}

async function uploadFixture(request, documentId, name, bytes, kind) {
  const attachmentId = createHash('sha256').update(bytes).digest('hex')
  const begin = await request('presentation-assets.v1', {
    operation: 'attachment_begin',
    documentId,
    attachmentId,
    name,
    sizeBytes: bytes.length,
    sha256: attachmentId,
  })
  if (
    begin?.attachmentId !== attachmentId ||
    begin.status !== 'uploading' ||
    begin.receivedBytes !== 0
  )
    throw new Error('PC smoke fixture already exists or upload failed')
  const chunk = await request('presentation-assets.v1', {
    operation: 'attachment_chunk',
    documentId,
    attachmentId,
    offset: 0,
    base64: bytes.toString('base64'),
  })
  if (
    chunk?.attachmentId !== attachmentId ||
    chunk.status !== 'uploading' ||
    chunk.receivedBytes !== bytes.length
  )
    throw new Error('PC smoke upload chunk failed')
  const finished = await request('presentation-assets.v1', {
    operation: 'attachment_finish',
    documentId,
    attachmentId,
  })
  if (
    finished?.attachmentId !== attachmentId ||
    finished.status !== 'ready' ||
    finished.kind !== kind ||
    finished.sha256 !== attachmentId
  )
    throw new Error('PC smoke attachment parse failed')
  return { attachmentId, metadata: finished }
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
  const fixtureIds = [options.textAttachmentId, options.imageAttachmentId]
  if (
    fixtureIds.some(
      (id) => id !== undefined && (typeof id !== 'string' || !/^[a-f0-9]{64}$/.test(id)),
    ) ||
    (fixtureIds[0] !== undefined && fixtureIds[0] === fixtureIds[1])
  )
    throw new Error('invalid smoke attachment fixture IDs')
  if (options.uploadFixtures && fixtureIds.some((id) => id !== undefined))
    throw new Error('upload smoke cannot use existing attachment IDs')
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
        capabilities: ['presentation.v1', 'presentation-assets.v1'],
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
      !approved.capabilities.includes('presentation.v1') ||
      !approved.capabilities.includes('presentation-assets.v1')
    )
      throw new Error('PC did not negotiate presentation.v1')
    async function request(capabilityName, body) {
      const requestId = randomUUID()
      socket.send(
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
      const start = await expected(next, 'relay.start', 30_000)
      if (
        start.request_id !== requestId ||
        start.session_id !== approved.session_id ||
        start.status !== 200 ||
        start.content_type !== 'application/json'
      )
        throw new Error('PC business request failed')
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
      try {
        return JSON.parse(Buffer.concat(chunks).toString('utf8'))
      } catch {
        throw new Error('invalid PC business JSON')
      }
    }
    const status = checkPcStatusResponse(
      await request('presentation.v1', { operation: 'status', documentId, projectId }),
      projectId,
    )
    const attachmentPageCount = checkAttachmentList(
      await request('presentation-assets.v1', { operation: 'attachment_list_assets', documentId }),
    )
    if (options.textAttachmentId) {
      const attachmentId = options.textAttachmentId
      const metadata = await request('presentation-assets.v1', {
        operation: 'attachment_metadata',
        documentId,
        attachmentId,
      })
      const read = await request('presentation-assets.v1', {
        operation: 'attachment_read',
        documentId,
        attachmentId,
        offset: 0,
        maxChars: 8000,
      })
      checkTextAttachment(metadata, read, attachmentId)
    }
    if (options.imageAttachmentId) {
      const attachmentId = options.imageAttachmentId
      const metadata = await request('presentation-assets.v1', {
        operation: 'attachment_metadata',
        documentId,
        attachmentId,
      })
      const asset = await request('presentation-assets.v1', {
        operation: 'attachment_asset',
        documentId,
        attachmentId,
      })
      checkImageAttachment(metadata, asset, attachmentId)
    }
    if (options.uploadFixtures) {
      const uploaded = []
      let failure
      try {
        const text = Buffer.from(`WisWork release smoke ${randomUUID()}\n`, 'utf8')
        const png = new PNG({ width: 1, height: 1 })
        png.data.set(Buffer.concat([randomBytes(3), Buffer.from([255])]))
        const image = PNG.sync.write(png)
        const textId = createHash('sha256').update(text).digest('hex')
        const imageId = createHash('sha256').update(image).digest('hex')
        for (const attachmentId of [textId, imageId]) {
          const existing = await request('presentation-assets.v1', {
            operation: 'attachment_metadata',
            documentId,
            attachmentId,
          })
          if (existing?.error !== 'not_found')
            throw new Error('PC smoke fixture already exists or metadata check failed')
        }
        // Once creation starts, cleanup is attempted even if a later upload step fails.
        uploaded.push(textId)
        const textResult = await uploadFixture(
          request,
          documentId,
          'wiswork-smoke.txt',
          text,
          'text',
        )
        const textRead = await request('presentation-assets.v1', {
          operation: 'attachment_read',
          documentId,
          attachmentId: textId,
          offset: 0,
          maxChars: 8000,
        })
        checkTextAttachment(textResult.metadata, textRead, textId)
        uploaded.push(imageId)
        const imageResult = await uploadFixture(
          request,
          documentId,
          'wiswork-smoke.png',
          image,
          'image',
        )
        const imageAsset = await request('presentation-assets.v1', {
          operation: 'attachment_asset',
          documentId,
          attachmentId: imageId,
        })
        checkImageAttachment(imageResult.metadata, imageAsset, imageId)
      } catch (error) {
        failure = error
      }
      for (const attachmentId of uploaded.reverse()) {
        try {
          const deleted = await request('presentation-assets.v1', {
            operation: 'attachment_delete',
            documentId,
            attachmentId,
          })
          if (deleted?.attachmentId !== attachmentId || deleted.deleted !== true)
            throw new Error('PC smoke fixture cleanup failed')
        } catch {
          failure = new Error(
            `PC smoke fixture cleanup failed for ${attachmentId}; inspect the test document`,
          )
        }
      }
      if (failure) throw failure
    }
    return {
      ...status,
      attachmentPageCount,
      textChecked: Boolean(options.textAttachmentId || options.uploadFixtures),
      imageChecked: Boolean(options.imageAttachmentId || options.uploadFixtures),
      uploadChecked: Boolean(options.uploadFixtures),
    }
  } finally {
    socket.terminate()
  }
}

if (process.argv[1] && fileURLToPath(import.meta.url) === process.argv[1]) {
  inspectPcBusiness(
    process.env.PPT_AGENT_SMOKE_RELAY_ORIGIN,
    process.env.PPT_AGENT_SMOKE_DOCUMENT_ID,
    process.env.PPT_AGENT_SMOKE_PROJECT_ID,
    {
      textAttachmentId: process.env.PPT_AGENT_SMOKE_TEXT_ATTACHMENT_ID,
      imageAttachmentId: process.env.PPT_AGENT_SMOKE_IMAGE_ATTACHMENT_ID,
      uploadFixtures: process.env.PPT_AGENT_SMOKE_UPLOAD === '1',
    },
  )
    .then((result) =>
      process.stdout.write(
        `Real PC business smoke passed: ${result.status}, ${result.slideCount} slides; ${result.attachmentPageCount} attachments on first page; text ${result.textChecked ? 'checked' : 'not configured'}, image ${result.imageChecked ? 'checked' : 'not configured'}, upload ${result.uploadChecked ? 'checked and cleaned' : 'not configured'}.\n`,
      ),
    )
    .catch((error) => {
      process.stderr.write(`${error.message}\n`)
      process.exitCode = 1
    })
}
