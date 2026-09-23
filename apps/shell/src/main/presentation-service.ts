import { createPresentationAttachmentService } from './presentation-attachments'
import {
  parsePresentationPlan,
  assertDeckMatchesPresentationPlan,
} from '@wiswork/pptx-engine/presentation-plan'
import { resolve } from 'node:path'
import { PresentationStore, assertPresentationId } from '@wiswork/project-store'
import {
  parsePresentationDeck,
  type PresentationCompileReport,
} from '@wiswork/pptx-engine/presentation'
import { compilePresentationDeck } from '@wiswork/pptx-engine/presentation-compiler'

const MAX_RESPONSE_BYTES = 15 * 1024 * 1024
const locks = new Map<string, Promise<void>>()
const errorCodes = new Set([
  'invalid_request',
  'invalid_plan',
  'plan_mismatch',
  'revision_conflict',
  'invalid_deck',
  'invalid_state',
  'document_mismatch',
  'request_conflict',
  'not_found',
  'aborted',
  'output_too_large',
  'unsupported_file',
  'attachment_conflict',
  'quota_exceeded',
  'digest_mismatch',
  'parse_failed',
])
const encode = (value: unknown): Uint8Array => Buffer.from(JSON.stringify(value), 'utf8')
function boundedResponse(value: unknown): Uint8Array {
  const response = encode(value)
  if (response.byteLength > MAX_RESPONSE_BYTES) throw new Error('output_too_large')
  return response
}
function savedDeck(value: unknown): ReturnType<typeof parsePresentationDeck> {
  try {
    return parsePresentationDeck(value)
  } catch {
    throw new Error('invalid_deck')
  }
}
function checkAbort(signal: AbortSignal): void {
  if (signal.aborted) throw new Error('aborted')
}

/** Main-process only. Compilation is durable; host import and visual verification are separate steps. */
export function createPresentationService(options: {
  userDataPath: string
  compile?: typeof compilePresentationDeck
}): (body: unknown, signal: AbortSignal) => Promise<Uint8Array> {
  const attachments = createPresentationAttachmentService(options)
  const store = new PresentationStore(options.userDataPath)
  const compile = options.compile ?? compilePresentationDeck
  return async (body, signal) => {
    try {
      checkAbort(signal)
      if (
        !body ||
        typeof body !== 'object' ||
        Array.isArray(body) ||
        Buffer.byteLength(JSON.stringify(body)) > 256 * 1024
      )
        throw new Error('invalid_request')
      const request = body as Record<string, unknown>
      if (
        [
          'attachment_begin',
          'attachment_chunk',
          'attachment_finish',
          'attachment_list',
          'attachment_read',
        ].includes(request.operation as string)
      )
        return boundedResponse(await attachments(request, signal))
      if (
        !['compile', 'get', 'status', 'resume', 'save_plan', 'get_plan'].includes(
          request.operation as string,
        )
      )
        throw new Error('invalid_request')
      const allowedKeys =
        request.operation === 'compile'
          ? ['operation', 'documentId', 'projectId', 'requestId', 'deck', 'planRevision']
          : request.operation === 'resume'
            ? ['operation', 'documentId', 'projectId', 'requestId']
            : request.operation === 'save_plan'
              ? ['operation', 'documentId', 'projectId', 'expectedRevision', 'plan']
              : ['operation', 'documentId', 'projectId']
      const requiredKeys = allowedKeys.filter(
        (key) => !(request.operation === 'compile' && ['projectId', 'planRevision'].includes(key)),
      )
      if (
        Object.keys(request).some((key) => !allowedKeys.includes(key)) ||
        requiredKeys.some((key) => !Object.hasOwn(request, key))
      )
        throw new Error('invalid_request')
      if (
        typeof request.documentId !== 'string' ||
        !request.documentId.trim() ||
        request.documentId.length > 2048
      )
        throw new Error('invalid_request')
      const documentId = request.documentId
      let deck: ReturnType<typeof parsePresentationDeck> | undefined
      if (request.operation === 'compile') {
        try {
          deck = parsePresentationDeck(request.deck)
        } catch {
          throw new Error('invalid_deck')
        }
        assertPresentationId(request.requestId)
      }
      if (request.operation === 'resume') assertPresentationId(request.requestId)
      if (
        request.operation === 'compile' &&
        request.planRevision !== undefined &&
        (!Number.isSafeInteger(request.planRevision) || Number(request.planRevision) < 1)
      )
        throw new Error('invalid_request')
      let plan: ReturnType<typeof parsePresentationPlan> | undefined
      if (request.operation === 'save_plan') {
        if (!Number.isSafeInteger(request.expectedRevision) || Number(request.expectedRevision) < 0)
          throw new Error('invalid_request')
        try {
          plan = parsePresentationPlan(request.plan)
        } catch {
          throw new Error('invalid_plan')
        }
      }
      const projectId = request.projectId ?? deck?.id
      assertPresentationId(projectId)
      if (plan && plan.projectId !== projectId) throw new Error('invalid_plan')
      if (deck && deck.id !== projectId) throw new Error('invalid_request')
      const key = `${resolve(options.userDataPath)}\0${projectId}`
      const previous = locks.get(key) ?? Promise.resolve()
      let release!: () => void
      const current = new Promise<void>((done) => {
        release = done
      })
      locks.set(key, current)
      await previous
      try {
        checkAbort(signal)
        if (request.operation === 'save_plan' || request.operation === 'get_plan') {
          const record =
            request.operation === 'save_plan'
              ? store.savePlan(projectId, documentId, request.expectedRevision as number, plan)
              : store.plan(projectId, documentId)
          if (!record) throw new Error('not_found')
          return boundedResponse({
            projectId,
            revision: record.revision,
            plan: parsePresentationPlan(record.plan),
          })
        }
        if (request.operation === 'status') {
          const savedPlan = store.plan(projectId, documentId)
          const plan = savedPlan
            ? { revision: savedPlan.revision, value: parsePresentationPlan(savedPlan.plan) }
            : undefined
          const history = store.history(projectId, documentId)
          const latest = history[0]
          if (!latest) {
            if (!plan) throw new Error('not_found')
            return boundedResponse({
              projectId,
              title: plan.value.title,
              status: 'planned',
              slideCount: plan.value.slides.length,
              slides: plan.value.slides.map(({ id, title }) => ({ id, title })),
              history: [],
              plan,
            })
          }
          const latestDeck = savedDeck(latest.deck)
          const compiled = store.latest(projectId, documentId)
          const checks =
            latest.status === 'compiled'
              ? (latest.result as { report: PresentationCompileReport }).report.checks
              : undefined
          return boundedResponse({
            projectId,
            title: latestDeck.title,
            status: latest.status,
            latestRequestId: latest.requestId,
            ...(plan ? { plan } : {}),
            ...(latest.plan ? { requestPlanRevision: latest.plan.revision } : {}),
            ...(compiled ? { latestCompiledRequestId: compiled.requestId } : {}),
            slideCount: latestDeck.slides.length,
            slides: latestDeck.slides.map(({ id, title }) => ({ id, title })),
            history: history.map((record) => ({
              requestId: record.requestId,
              sequence: record.sequence,
              status: record.status,
              slideCount: savedDeck(record.deck).slides.length,
            })),
            ...(checks
              ? {
                  checks: {
                    structure: checks.structure,
                    geometry: checks.geometry,
                    render: checks.render,
                    sources: checks.sources,
                    roundTrip: checks.roundTrip,
                  },
                }
              : {}),
          })
        }
        if (request.operation === 'get') {
          const record = store.latest(projectId, documentId)
          if (!record) throw new Error('not_found')
          return boundedResponse(record.result)
        }
        const requestId = request.requestId as string
        let record = store.request(projectId, documentId, requestId)
        if (request.operation === 'compile') {
          const saved = record ? undefined : store.plan(projectId, documentId)
          const binding =
            record?.plan ?? (saved ? { revision: saved.revision, plan: saved.plan } : undefined)
          if (request.planRevision !== binding?.revision)
            throw new Error(record ? 'request_conflict' : 'revision_conflict')
          if (binding) {
            try {
              assertDeckMatchesPresentationPlan(deck!, parsePresentationPlan(binding.plan))
            } catch {
              throw new Error(record ? 'request_conflict' : 'plan_mismatch')
            }
          }
          record = store.begin(projectId, documentId, requestId, deck, binding)
        }
        if (!record) throw new Error('not_found')
        if (record.status === 'compiled') return boundedResponse(record.result)
        const inputDeck = savedDeck(record.deck)
        if (inputDeck.id !== projectId) throw new Error('invalid_deck')
        if (record.plan) {
          try {
            assertDeckMatchesPresentationPlan(inputDeck, parsePresentationPlan(record.plan.plan))
          } catch {
            throw new Error('plan_mismatch')
          }
        }
        const compiled = await compile(inputDeck)
        checkAbort(signal)
        if (compiled.bytes.byteLength > 10 * 1024 * 1024) throw new Error('output_too_large')
        const result = {
          projectId,
          requestId,
          status: 'compiled',
          pptxBase64: Buffer.from(compiled.bytes).toString('base64'),
          report: compiled.report,
        }
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
