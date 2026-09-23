import { resolve } from 'node:path'
import { PresentationStore, assertPresentationId } from '@wiswork/project-store'
import { parsePresentationDeck } from '@wiswork/pptx-engine/presentation'
import { compilePresentationDeck } from '@wiswork/pptx-engine/presentation-compiler'

const MAX_RESPONSE_BYTES = 15 * 1024 * 1024
const locks = new Map<string, Promise<void>>()
const errorCodes = new Set(['invalid_request', 'invalid_deck', 'invalid_state', 'document_mismatch', 'request_conflict', 'not_found', 'aborted', 'output_too_large'])
const encode = (value: unknown): Uint8Array => Buffer.from(JSON.stringify(value), 'utf8')
function checkAbort(signal: AbortSignal): void { if (signal.aborted) throw new Error('aborted') }

/** Main-process only. Compilation is durable; host import and visual verification are separate steps. */
export function createPresentationService(options: {
  userDataPath: string
  compile?: typeof compilePresentationDeck
}): (body: unknown, signal: AbortSignal) => Promise<Uint8Array> {
  const store = new PresentationStore(options.userDataPath)
  const compile = options.compile ?? compilePresentationDeck
  return async (body, signal) => {
    try {
      checkAbort(signal)
      if (!body || typeof body !== 'object' || Array.isArray(body) || Buffer.byteLength(JSON.stringify(body)) > 256 * 1024) throw new Error('invalid_request')
      const request = body as Record<string, unknown>
      if (request.operation !== 'compile' && request.operation !== 'get') throw new Error('invalid_request')
      const allowedKeys = request.operation === 'compile'
        ? ['operation', 'documentId', 'projectId', 'requestId', 'deck']
        : ['operation', 'documentId', 'projectId']
      if (Object.keys(request).some(key => !allowedKeys.includes(key))) throw new Error('invalid_request')
      if (typeof request.documentId !== 'string' || !request.documentId.trim() || request.documentId.length > 2048) throw new Error('invalid_request')
      const documentId = request.documentId
      let deck: ReturnType<typeof parsePresentationDeck> | undefined
      if (request.operation === 'compile') {
        try { deck = parsePresentationDeck(request.deck) } catch { throw new Error('invalid_deck') }
        assertPresentationId(request.requestId)
      }
      const projectId = request.projectId ?? deck?.id
      assertPresentationId(projectId)
      if (deck && deck.id !== projectId) throw new Error('invalid_request')
      const key = `${resolve(options.userDataPath)}\0${projectId}`
      const previous = locks.get(key) ?? Promise.resolve()
      let release!: () => void
      const current = new Promise<void>(done => { release = done })
      locks.set(key, current)
      await previous
      try {
        checkAbort(signal)
        if (request.operation === 'get') {
          const record = store.latest(projectId, documentId)
          if (!record) throw new Error('not_found')
          const response = encode(record.result)
          if (response.byteLength > MAX_RESPONSE_BYTES) throw new Error('output_too_large')
          return response
        }
        const requestId = request.requestId as string
        const record = store.begin(projectId, documentId, requestId, deck)
        if (record.status === 'compiled') return encode(record.result)
        const compiled = await compile(deck)
        checkAbort(signal)
        if (compiled.bytes.byteLength > 10 * 1024 * 1024) throw new Error('output_too_large')
        const result = { projectId, requestId, status: 'compiled', pptxBase64: Buffer.from(compiled.bytes).toString('base64'), report: compiled.report }
        const response = encode(result)
        if (response.byteLength > MAX_RESPONSE_BYTES) throw new Error('output_too_large')
        store.complete(record, result)
        return response
      } finally {
        release()
        if (locks.get(key) === current) locks.delete(key)
      }
    } catch (error) {
      const message = error instanceof Error ? error.message : ''
      return encode({ error: errorCodes.has(message) ? message : 'compile_failed' })
    }
  }
}
