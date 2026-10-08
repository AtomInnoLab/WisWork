import { createHash, randomBytes, randomUUID } from 'node:crypto'
import { fileURLToPath } from 'node:url'
import { PDFDocument } from 'pdf-lib'
import JSZip from 'jszip'
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
    if (waiter) {
      if (error) waiter.reject(error)
      else waiter.resolve(frame)
    } else if (error) failure = error
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

export async function checkCompiledDelivery(compiled, exported, projectId, expectedTexts) {
  const pageCount = expectedTexts?.length
  if (
    !Array.isArray(expectedTexts) ||
    pageCount < 1 ||
    pageCount > 8 ||
    compiled?.projectId !== projectId ||
    compiled.status !== 'compiled' ||
    compiled.report?.slideCount !== pageCount ||
    typeof compiled.pptxBase64 !== 'string' ||
    exported?.projectId !== projectId ||
    exported.status !== 'exported' ||
    exported.source !== 'compiled' ||
    exported.requestId !== compiled.requestId ||
    exported.slideCount !== pageCount ||
    typeof exported.pdfBase64 !== 'string'
  )
    throw new Error('PC compiled delivery response invalid')
  const pptx = Buffer.from(compiled.pptxBase64, 'base64')
  const pdf = Buffer.from(exported.pdfBase64, 'base64')
  if (
    !pptx.length ||
    pptx.length > 10 * 1024 * 1024 ||
    pptx.toString('base64') !== compiled.pptxBase64 ||
    !pdf.length ||
    pdf.length > 10 * 1024 * 1024 ||
    pdf.toString('base64') !== exported.pdfBase64 ||
    !pdf.subarray(0, 5).equals(Buffer.from('%PDF-'))
  )
    throw new Error('PC compiled delivery bytes invalid')
  const zip = await JSZip.loadAsync(pptx)
  for (let i = 0; i < pageCount; i++) {
    const slide = await zip.file(`ppt/slides/slide${i + 1}.xml`)?.async('string')
    if (!slide?.includes(`<a:t>${expectedTexts[i]}</a:t>`))
      throw new Error('PC compiled delivery content invalid')
  }
  if (
    zip.file(`ppt/slides/slide${pageCount + 1}.xml`) ||
    (await PDFDocument.load(pdf)).getPageCount() !== pageCount
  )
    throw new Error('PC compiled delivery content invalid')
  return { pptxSha256: createHash('sha256').update(pptx).digest('hex'), pdfBytes: pdf.length }
}

export function releaseProductionFixture(projectId) {
  if (typeof projectId !== 'string' || !/^[A-Za-z0-9_-]{1,128}$/.test(projectId))
    throw new Error('invalid release smoke project ID')
  const expectedSlideTexts = Array.from(
    { length: 8 },
    (_, index) => `Release smoke page ${index + 1}`,
  )
  const style = {
    fontFace: 'Arial',
    background: 'FFFFFF',
    textColor: '111111',
    accentColor: '3366FF',
  }
  const slides = expectedSlideTexts.map((text, index) => ({
    id: `slide-${index + 1}`,
    title: `Page ${index + 1}`,
    elements: [{ id: 'title', kind: 'text', text, x: 1, y: 1, w: 8, h: 1 }],
  }))
  slides[2].elements.push({
    id: 'image',
    kind: 'image',
    x: 1,
    y: 2.5,
    w: 3,
    h: 3,
    assetId: 'smoke-image',
    fit: 'contain',
  })
  slides[3].elements.push({
    id: 'shape',
    kind: 'shape',
    x: 1,
    y: 2.5,
    w: 3,
    h: 2,
    shape: 'roundRect',
  })
  slides[5].elements.push({
    id: 'table',
    kind: 'table',
    x: 1,
    y: 2.5,
    w: 8,
    h: 2,
    rows: [
      ['Plan', 'Result'],
      ['A', '120'],
      ['B', '90'],
    ],
  })
  slides[6].elements.push({
    id: 'chart',
    kind: 'chart',
    x: 1,
    y: 2.5,
    w: 8,
    h: 3,
    chartType: 'bar',
    categories: ['A', 'B'],
    series: [{ name: 'Result', values: [120, 90] }],
  })
  const deck = {
    version: 1,
    id: projectId,
    title: 'Release smoke',
    style,
    assets: [
      {
        id: 'smoke-image',
        mime: 'image/png',
        width: 1,
        height: 1,
        base64:
          'iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAYAAAAfFcSJAAAAC0lEQVR4AWP4DwQACfsD/c8LaHIAAAAASUVORK5CYII=',
        source: 'Synthetic release fixture',
      },
    ],
    claims: [],
    slides,
  }
  const plan = {
    version: 1,
    projectId,
    title: deck.title,
    brief: {
      objective: 'Verify an eight-page editable release delivery',
      audience: 'Release test',
      language: 'en-US',
      minutes: 8,
      requiredContent: [],
      constraints: [],
    },
    sources: [],
    claims: [],
    style,
    slides: slides.map((slide) => ({
      id: slide.id,
      title: slide.title,
      purpose: 'Verify native text',
      claimIds: [],
      layout: 'content',
      requiredAssets: slide.id === 'slide-3' ? ['smoke-image'] : [],
      acceptanceCriteria: ['Native text remains editable'],
    })),
  }
  return {
    requestId: `release-production-${randomBytes(8).toString('hex')}`,
    deck,
    plan,
    expectedSlideTexts,
  }
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
  for (let offset = 0; offset < bytes.length; offset += 64 * 1024) {
    const part = bytes.subarray(offset, offset + 64 * 1024)
    const chunk = await request('presentation-assets.v1', {
      operation: 'attachment_chunk',
      documentId,
      attachmentId,
      offset,
      base64: part.toString('base64'),
    })
    if (
      chunk?.attachmentId !== attachmentId ||
      chunk.status !== 'uploading' ||
      chunk.receivedBytes !== offset + part.length
    )
      throw new Error('PC smoke upload chunk failed')
  }
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
    throw new Error(
      `PC smoke ${kind} attachment parse failed: ${finished?.error ?? finished?.status ?? 'invalid_response'}`,
    )
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
  const fallback = options.remoteImageCandidates
  if (
    fallback &&
    (!Array.isArray(fallback.urls) ||
      fallback.urls.length !== 2 ||
      fallback.urls.some((url) => typeof url !== 'string' || !url.startsWith('https://')) ||
      fallback.urls[0] === fallback.urls[1] ||
      !/^[a-f0-9]{64}$/.test(fallback.attachmentId))
  )
    throw new Error('invalid remote image fallback fixture')
  if (
    [
      options.readExistingProduction,
      options.beginProductionOnly,
      options.runExistingProduction,
    ].filter(Boolean).length > 1 ||
    (options.createProduction && !options.productionFixture) ||
    (options.createProduction &&
      (options.readExistingProduction ||
        options.beginProductionOnly ||
        options.runExistingProduction)) ||
    ((options.readExistingProduction ||
      options.beginProductionOnly ||
      options.runExistingProduction) &&
      !options.productionFixture)
  )
    throw new Error('existing production smoke requires a production fixture')
  if (
    options.compiledRequestId !== undefined &&
    (typeof options.compiledRequestId !== 'string' ||
      !/^[A-Za-z0-9_-]{1,128}$/.test(options.compiledRequestId) ||
      !Array.isArray(options.expectedSlideTexts) ||
      options.expectedSlideTexts.length < 1 ||
      options.expectedSlideTexts.length > 8 ||
      options.expectedSlideTexts.some(
        (value) => typeof value !== 'string' || !/^[A-Za-z0-9 _-]{1,128}$/.test(value),
      ))
  )
    throw new Error('invalid compiled delivery smoke fixture')
  if (
    options.productionFixture &&
    (typeof options.productionFixture.requestId !== 'string' ||
      !/^[A-Za-z0-9_-]{1,128}$/.test(options.productionFixture.requestId) ||
      options.productionFixture.deck?.id !== projectId ||
      options.productionFixture.plan?.projectId !== projectId ||
      !Array.isArray(options.productionFixture.expectedSlideTexts) ||
      options.productionFixture.expectedSlideTexts.length !==
        options.productionFixture.deck?.slides?.length ||
      options.productionFixture.expectedSlideTexts.length < 1 ||
      options.productionFixture.expectedSlideTexts.length > 8)
  )
    throw new Error('invalid production smoke fixture')
  if (
    fallback &&
    options.productionFixture &&
    !options.productionFixture.deck.assets?.some(
      (asset) => asset.attachmentId === fallback.attachmentId,
    )
  )
    throw new Error('fallback image must be used by production fixture')
  if (
    options.derivedPageFixture &&
    (!options.createProduction ||
      !options.productionFixture ||
      typeof options.derivedPageFixture.requestId !== 'string' ||
      !/^[A-Za-z0-9_-]{1,128}$/.test(options.derivedPageFixture.requestId) ||
      options.derivedPageFixture.requestId === options.productionFixture.requestId ||
      typeof options.derivedPageFixture.pageId !== 'string' ||
      options.derivedPageFixture.slide?.id !== options.derivedPageFixture.pageId ||
      !options.productionFixture.deck.slides.some(
        (slide) => slide.id === options.derivedPageFixture.pageId,
      ))
  )
    throw new Error('invalid derived page smoke fixture')
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
        capabilities: [
          'presentation.v1',
          'presentation-assets.v1',
          ...(fallback ? ['presentation-remote-images.v1'] : []),
          ...(options.compiledRequestId ? ['presentation-pdf.v1'] : []),
          ...(options.productionFixture ? ['presentation-production-pdf.v1'] : []),
        ],
      }),
    )
    const created = await expected(next, 'office.created', 10_000)
    if (!/^[0-9]{6}$/.test(created.verification_code))
      throw new Error('invalid relay pairing invitation')
    options.onProgress?.('invited')
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
      !approved.capabilities.includes('presentation-assets.v1') ||
      (fallback && !approved.capabilities.includes('presentation-remote-images.v1')) ||
      (options.compiledRequestId && !approved.capabilities.includes('presentation-pdf.v1')) ||
      (options.productionFixture &&
        !approved.capabilities.includes('presentation-production-pdf.v1'))
    )
      throw new Error('PC did not negotiate required presentation capabilities')
    options.onProgress?.('paired')
    async function nextResponse(timeout) {
      const deadline = Date.now() + timeout
      for (;;) {
        const remaining = deadline - Date.now()
        if (remaining <= 0) throw new Error('relay response timed out')
        const frame = await next(remaining)
        if (frame.type !== 'relay.session_state') return frame
        if (frame.session_id !== approved.session_id)
          throw new Error('session state identity mismatch')
      }
    }
    async function request(capabilityName, body) {
      const requestId = randomUUID()
      options.onProgress?.(`request:${body.operation}`)
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
      const start = await expected(nextResponse, 'relay.start', Math.max(30_000, timeoutMs))
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
        const frame = await nextResponse(Math.max(30_000, timeoutMs))
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
    const sourceAttachments =
      options.sourceAttachments ?? (options.sourceAttachment ? [options.sourceAttachment] : [])
    let remoteImageFallback
    if (fallback) {
      const importUrl = (url) =>
        request('presentation-remote-images.v1', {
          operation: 'attachment_import_url',
          documentId,
          url,
        })
      const failed = await importUrl(fallback.urls[0])
      if (failed?.error !== 'remote_image_unavailable')
        throw new Error('PC primary image did not fail as expected')
      const primaryFailureAt = new Date().toISOString()
      const imported = await importUrl(fallback.urls[1])
      if (
        imported?.status !== 'ready' ||
        imported.kind !== 'image' ||
        imported.attachmentId !== fallback.attachmentId ||
        imported.source !== fallback.urls[1]
      )
        throw new Error('PC fallback image import invalid')
      const cached = await importUrl(fallback.urls[1])
      if (cached?.attachmentId !== imported.attachmentId || cached.status !== 'ready')
        throw new Error('PC fallback image cache invalid')
      const asset = await request('presentation-assets.v1', {
        operation: 'attachment_asset',
        documentId,
        attachmentId: imported.attachmentId,
      })
      checkImageAttachment(imported, asset, imported.attachmentId)
      remoteImageFallback = {
        attachmentId: imported.attachmentId,
        assetSha256: imported.assetSha256,
        attemptedUrls: fallback.urls,
        primaryFailureAt,
        cacheHit: true,
        projectId,
        productionRequestId: options.productionFixture?.requestId,
      }
    }
    if (sourceAttachments.length) {
      const sourceHashes = new Set(sourceAttachments.map((source) => source.sha256))
      if (
        !options.productionFixture?.plan?.sources?.length ||
        sourceHashes.size !== sourceAttachments.length ||
        options.productionFixture.plan.sources.some(
          (item) =>
            !sourceHashes.has(
              item.snapshotAttachmentId ?? /^attachment:([a-f0-9]{64})$/.exec(item.uri)?.[1],
            ),
        )
      )
        throw new Error('invalid source-backed production fixture')
      for (const source of sourceAttachments) {
        if (
          typeof source.name !== 'string' ||
          !/\.(?:pdf|csv|json|docx|txt|md|html|xlsx|pptx)$/i.test(source.name) ||
          !Buffer.isBuffer(source.bytes) ||
          source.bytes.length < 1 ||
          source.bytes.length > 50 * 1024 * 1024 ||
          createHash('sha256').update(source.bytes).digest('hex') !== source.sha256
        )
          throw new Error('invalid source-backed production fixture')
        await uploadFixture(request, documentId, source.name, source.bytes, 'text')
      }
    }
    if (options.createProduction) {
      const fixture = options.productionFixture
      const saved = await request('presentation.v1', {
        operation: 'save_plan',
        documentId,
        projectId,
        expectedRevision: 0,
        plan: fixture.plan,
      })
      if (saved?.revision !== 1) throw new Error('PC production plan was not saved')
    }
    const status = checkPcStatusResponse(
      await request('presentation.v1', { operation: 'status', documentId, projectId }),
      projectId,
    )
    let manualObservation
    if (options.manualObservation) {
      const observationId = 'electron_manual_observation'
      const base = { documentId, projectId, observationId }
      if (options.manualObservation === 'create') {
        const shape = {
          id: 'smoke-shape',
          name: 'Smoke text',
          type: 'TextBox',
          left: 1,
          top: 1,
          width: 4,
          height: 1,
          text: 'Before edit',
        }
        const begun = await request('presentation.v1', {
          operation: 'manual_observation_begin',
          ...base,
          slideId: 'smoke-slide',
          shape,
        })
        const digest = begun?.observation?.before?.digest
        if (!/^[a-f0-9]{64}$/.test(digest)) throw new Error('PC manual observation begin failed')
        const completed = await request('presentation.v1', {
          operation: 'manual_observation_complete',
          ...base,
          expectedBeforeDigest: digest,
          shape: { ...shape, text: 'After edit' },
        })
        manualObservation = completed?.observation
        if (manualObservation?.after?.digest === digest || !manualObservation?.after?.digest)
          throw new Error('PC manual observation complete failed')
      } else if (options.manualObservation === 'read') {
        manualObservation = (
          await request('presentation.v1', {
            operation: 'manual_observation_get',
            ...base,
          })
        )?.observation
        if (manualObservation?.after?.shape?.text !== 'After edit')
          throw new Error('PC manual observation recovery failed')
      } else throw new Error('invalid manual observation smoke mode')
    }
    let compiledDelivery
    if (options.compiledRequestId) {
      const compiled = await request('presentation.v1', {
        operation: 'get',
        documentId,
        projectId,
      })
      if (compiled?.requestId !== options.compiledRequestId)
        throw new Error('PC compiled request identity mismatch')
      const exported = await request('presentation-pdf.v1', {
        operation: 'export_pdf',
        documentId,
        projectId,
        requestId: options.compiledRequestId,
        source: 'compiled',
      })
      compiledDelivery = await checkCompiledDelivery(
        compiled,
        exported,
        projectId,
        options.expectedSlideTexts,
      )
      const source = await request('presentation.v1', {
        operation: 'read_import_source',
        documentId,
        projectId,
        requestId: options.compiledRequestId,
        source: 'compiled',
      })
      if (
        source?.version !== 1 ||
        source.documentId !== documentId ||
        source.projectId !== projectId ||
        source.requestId !== options.compiledRequestId ||
        source.source !== 'compiled' ||
        source.artifactDigest !== createHash('sha256').update(compiled.pptxBase64).digest('hex') ||
        !Array.isArray(source.pages) ||
        source.pages.length !== options.expectedSlideTexts.length ||
        source.pages.some(
          (page, index) =>
            page.id !== compiled.pages?.[index]?.id ||
            page.title !== compiled.pages?.[index]?.title ||
            page.sourceSlideId !== compiled.pages?.[index]?.sourceSlideId,
        )
      )
        throw new Error('PC compiled import source invalid')
    }
    let productionDelivery
    if (options.productionFixture) {
      const fixture = options.productionFixture
      const base = { documentId, projectId, requestId: fixture.requestId }
      let produced
      if (options.readExistingProduction) {
        produced = await request('presentation.v1', { operation: 'production_status', ...base })
      } else if (options.runExistingProduction) {
        produced = await request('presentation.v1', { operation: 'production_run', ...base })
      } else {
        if (!options.createProduction) {
          const saved = await request('presentation.v1', {
            operation: 'save_plan',
            documentId,
            projectId,
            expectedRevision: 0,
            plan: fixture.plan,
          })
          if (saved?.revision !== 1) throw new Error('PC production plan was not saved')
        }
        const begun = await request('presentation.v1', {
          operation: 'production_begin',
          ...base,
          planRevision: 1,
          deck: fixture.deck,
        })
        if (begun?.status !== 'pending' || begun.total !== fixture.expectedSlideTexts.length)
          throw new Error('PC page production did not begin')
        if (options.beginProductionOnly)
          return { ...status, productionDelivery: { status: 'pending', total: begun.total } }
        produced = await request('presentation.v1', { operation: 'production_run', ...base })
      }
      if (
        produced?.status !== 'compiled' ||
        produced.compiledCount !== fixture.expectedSlideTexts.length ||
        produced.total !== fixture.expectedSlideTexts.length ||
        !Array.isArray(produced.pages) ||
        produced.pages.length !== fixture.expectedSlideTexts.length ||
        produced.pages.some(
          (page, index) => page.id !== fixture.deck.slides[index]?.id || page.state !== 'compiled',
        )
      )
        throw new Error(
          `PC page production incomplete: ${JSON.stringify({
            status: produced?.status,
            error: produced?.error,
            compiledCount: produced?.compiledCount,
            pages: produced?.pages?.map((page) => ({
              id: page.id,
              state: page.state,
              error: page.error,
            })),
          })}`,
        )
      const pageDigests = []
      const pagePptxBase64 = []
      const sourcePages = []
      for (const [index, slide] of fixture.deck.slides.entries()) {
        const page = await request('presentation.v1', {
          operation: 'production_page',
          ...base,
          pageId: slide.id,
        })
        if (
          page?.projectId !== projectId ||
          page.requestId !== fixture.requestId ||
          page.pageId !== slide.id ||
          page.status !== 'compiled' ||
          page.planRevision !== 1 ||
          page.report?.slideCount !== 1 ||
          typeof page.sourceSlideId !== 'string' ||
          typeof page.pptxBase64 !== 'string'
        )
          throw new Error('PC production page identity invalid')
        if (
          options.expectedFontResolution &&
          (page.report?.fontResolution?.requested !== options.expectedFontResolution.requested ||
            page.report.fontResolution.used !== options.expectedFontResolution.used ||
            page.report.fontResolution.substituted !== options.expectedFontResolution.substituted)
        )
          throw new Error('PC production font fallback mismatch')
        const bytes = Buffer.from(page.pptxBase64, 'base64')
        if (
          !bytes.length ||
          bytes.length > 10 * 1024 * 1024 ||
          bytes.toString('base64') !== page.pptxBase64
        )
          throw new Error('PC production page bytes invalid')
        const zip = await JSZip.loadAsync(bytes)
        const xml = await zip.file('ppt/slides/slide1.xml')?.async('string')
        if (
          !xml?.includes(`<a:t>${fixture.expectedSlideTexts[index]}</a:t>`) ||
          zip.file('ppt/slides/slide2.xml')
        )
          throw new Error('PC production page content invalid')
        const images = slide.elements.filter((element) => element.kind === 'image')
        if (images.length) {
          const media = Object.keys(zip.files).filter((path) =>
            /^ppt\/media\/[^/]+\.png$/.test(path),
          )
          if (
            (xml.match(/<p:pic>/g) ?? []).length !== images.length ||
            media.length !== images.length
          )
            throw new Error('PC production native image missing')
          const expected = images.map((element) =>
            fixture.deck.assets.find((asset) => asset.id === element.assetId),
          )
          if (
            expected.some(
              (asset) => !asset?.base64 && asset?.attachmentId !== fallback?.attachmentId,
            )
          )
            throw new Error('PC production native image missing')
          const actualDigests = await Promise.all(
            media.map(async (path) => {
              const image = await zip.file(path).async('nodebuffer')
              PNG.sync.read(image)
              return createHash('sha256').update(image).digest('hex')
            }),
          )
          const expectedDigests = expected.map((asset) =>
            asset.base64
              ? createHash('sha256').update(Buffer.from(asset.base64, 'base64')).digest('hex')
              : remoteImageFallback?.assetSha256,
          )
          if (actualDigests.sort().join(',') !== expectedDigests.sort().join(','))
            throw new Error('PC production native image changed')
        }
        if (
          slide.elements.some(
            (element) => element.kind === 'shape' && element.shape === 'roundRect',
          ) &&
          !xml.includes('prst="roundRect"')
        )
          throw new Error('PC production native shape missing')
        if (slide.elements.some((element) => element.kind === 'table') && !xml.includes('<a:tbl>'))
          throw new Error('PC production native table missing')
        if (slide.elements.some((element) => element.kind === 'chart')) {
          const chart = Object.keys(zip.files).filter((path) =>
            /^ppt\/charts\/chart\d+\.xml$/.test(path),
          )
          const workbook = Object.keys(zip.files).filter((path) =>
            /^ppt\/embeddings\/[^/]+\.xlsx$/.test(path),
          )
          if (!xml.includes('<c:chart ') || chart.length !== 1 || workbook.length !== 1)
            throw new Error('PC production native chart missing')
          const chartXml = await zip.file(chart[0]).async('string')
          if (!chartXml.includes('<c:ser>') || !chartXml.includes('<c:externalData '))
            throw new Error('PC production native chart data missing')
          const book = await JSZip.loadAsync(await zip.file(workbook[0]).async('nodebuffer'))
          if (!book.file('xl/workbook.xml') || !book.file('xl/worksheets/sheet1.xml'))
            throw new Error('PC production native chart workbook missing')
          const sheet = await book.file('xl/worksheets/sheet1.xml').async('string')
          for (const series of slide.elements.find((element) => element.kind === 'chart').series)
            for (const value of series.values)
              if (!chartXml.includes(`<c:v>${value}</c:v>`) || !sheet.includes(`<v>${value}</v>`))
                throw new Error('PC production native chart data changed')
        }
        pageDigests.push(createHash('sha256').update(bytes).digest('hex'))
        pagePptxBase64.push(page.pptxBase64)
        sourcePages.push({ id: slide.id, title: slide.title, sourceSlideId: page.sourceSlideId })
      }
      const source = await request('presentation.v1', {
        operation: 'read_import_source',
        ...base,
        source: 'production',
      })
      const importContent = JSON.stringify({
        documentId,
        projectId,
        requestId: fixture.requestId,
        planRevision: 1,
        pages: sourcePages,
        pagePptxBase64,
      })
      if (
        source?.version !== 1 ||
        source.documentId !== documentId ||
        source.projectId !== projectId ||
        source.requestId !== fixture.requestId ||
        source.source !== 'production' ||
        source.planRevision !== 1 ||
        source.artifactDigest !== createHash('sha256').update(importContent).digest('hex') ||
        !Array.isArray(source.pages) ||
        source.pages.length !== sourcePages.length ||
        source.pages.some(
          (page, index) =>
            page.id !== sourcePages[index].id ||
            page.title !== sourcePages[index].title ||
            page.sourceSlideId !== sourcePages[index].sourceSlideId,
        )
      )
        throw new Error('PC production import source invalid')
      const exported = await request('presentation-production-pdf.v1', {
        operation: 'export_pdf',
        ...base,
        source: 'production',
      })
      const pdf = Buffer.from(exported?.pdfBase64 ?? '', 'base64')
      if (
        exported?.projectId !== projectId ||
        exported.requestId !== fixture.requestId ||
        exported.source !== 'production' ||
        exported.status !== 'exported' ||
        exported.slideCount !== fixture.expectedSlideTexts.length ||
        pdf.length < 100 ||
        pdf.length > 10 * 1024 * 1024 ||
        pdf.toString('base64') !== exported.pdfBase64 ||
        (await PDFDocument.load(pdf)).getPageCount() !== fixture.expectedSlideTexts.length
      )
        throw new Error('PC production PDF invalid')
      let derivedPage
      if (options.derivedPageFixture) {
        const revision = options.derivedPageFixture
        const childBase = { documentId, projectId, requestId: revision.requestId }
        const started = await request('presentation.v1', {
          operation: 'production_rebuild_page',
          ...childBase,
          parentRequestId: fixture.requestId,
          pageId: revision.pageId,
          slide: revision.slide,
        })
        if (
          started?.compiledCount !== fixture.deck.slides.length - 1 ||
          started?.revision?.parentRequestId !== fixture.requestId ||
          started.revision.pageId !== revision.pageId ||
          !/^[a-f0-9]{64}$/.test(started.revision.parentInputDigest)
        )
          throw new Error('PC derived page parent binding invalid')
        const completed = await request('presentation.v1', {
          operation: 'production_run',
          ...childBase,
        })
        if (
          completed?.status !== 'compiled' ||
          completed.compiledCount !== fixture.deck.slides.length ||
          completed.revision?.parentRequestId !== fixture.requestId ||
          completed.revision.pageId !== revision.pageId
        )
          throw new Error('PC derived page production incomplete')
        const childDigests = []
        for (const [index, slide] of fixture.deck.slides.entries()) {
          const page = await request('presentation.v1', {
            operation: 'production_page',
            ...childBase,
            pageId: slide.id,
          })
          if (
            page?.status !== 'compiled' ||
            page.pageId !== slide.id ||
            typeof page.pptxBase64 !== 'string'
          )
            throw new Error('PC derived page receipt invalid')
          const childBytes = Buffer.from(page.pptxBase64, 'base64')
          const changed = slide.id === revision.pageId
          if (changed === (page.pptxBase64 === pagePptxBase64[index]))
            throw new Error('PC derived page preservation mismatch')
          if (changed) {
            const zip = await JSZip.loadAsync(childBytes)
            const xml = await zip.file('ppt/slides/slide1.xml')?.async('string')
            if (!xml || (xml.match(/prst="roundRect"/g) ?? []).length !== 3)
              throw new Error('PC derived page native flow missing')
          }
          childDigests.push(createHash('sha256').update(childBytes).digest('hex'))
        }
        const parentTarget = await request('presentation.v1', {
          operation: 'production_page',
          ...base,
          pageId: revision.pageId,
        })
        const targetIndex = fixture.deck.slides.findIndex((slide) => slide.id === revision.pageId)
        if (parentTarget?.pptxBase64 !== pagePptxBase64[targetIndex])
          throw new Error('PC derived page overwrote parent')
        derivedPage = {
          requestId: revision.requestId,
          parentRequestId: fixture.requestId,
          pageId: revision.pageId,
          parentInputDigest: started.revision.parentInputDigest,
          pageDigests: childDigests,
        }
      }
      productionDelivery = {
        pageDigests,
        pdfBytes: pdf.length,
        ...(derivedPage ? { derivedPage } : {}),
      }
    }
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
      ...(compiledDelivery ? { compiledDelivery } : {}),
      ...(productionDelivery ? { productionDelivery } : {}),
      ...(sourceAttachments.length ? { sourceChecked: true } : {}),
      ...(remoteImageFallback ? { remoteImageFallback } : {}),
      ...(manualObservation ? { manualObservation } : {}),
    }
  } finally {
    socket.terminate()
  }
}

if (process.argv[1] && fileURLToPath(import.meta.url) === process.argv[1]) {
  const production = process.env.PPT_AGENT_SMOKE_PRODUCTION === '1'
  const projectPrefix = process.env.PPT_AGENT_SMOKE_PROJECT_ID
  const projectId = production
    ? `${(projectPrefix ?? '').slice(0, 100)}-${randomBytes(6).toString('hex')}`
    : projectPrefix
  if (production && !/^[A-Za-z0-9_-]{1,128}$/.test(projectPrefix ?? '')) {
    process.stderr.write('Production smoke requires a valid dedicated test project prefix.\n')
    process.exitCode = 1
  } else
    inspectPcBusiness(
      process.env.PPT_AGENT_SMOKE_RELAY_ORIGIN,
      process.env.PPT_AGENT_SMOKE_DOCUMENT_ID,
      projectId,
      {
        textAttachmentId: process.env.PPT_AGENT_SMOKE_TEXT_ATTACHMENT_ID,
        imageAttachmentId: process.env.PPT_AGENT_SMOKE_IMAGE_ATTACHMENT_ID,
        uploadFixtures: process.env.PPT_AGENT_SMOKE_UPLOAD === '1',
        ...(production
          ? { productionFixture: releaseProductionFixture(projectId), createProduction: true }
          : {}),
      },
    )
      .then((result) =>
        process.stdout.write(
          `Real PC business smoke passed: project ${result.projectId}, ${result.status}, ${result.slideCount} slides; production ${result.productionDelivery?.pageDigests.length ?? 'not configured'} pages; ${result.attachmentPageCount} attachments on first page; text ${result.textChecked ? 'checked' : 'not configured'}, image ${result.imageChecked ? 'checked' : 'not configured'}, upload ${result.uploadChecked ? 'checked and cleaned' : 'not configured'}.\n`,
        ),
      )
      .catch((error) => {
        process.stderr.write(`${error.message}\n`)
        process.exitCode = 1
      })
}
