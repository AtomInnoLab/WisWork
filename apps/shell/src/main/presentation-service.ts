import {
  activePresentationRequest,
  handlePresentationJob,
  hasPresentationWorker,
  presentationJobOperations,
} from './presentation-jobs'
import { parsePresentationClaimReview } from '@wiswork/pptx-engine/presentation-claim-review'
import { createPresentationPageBackupService } from './presentation-page-backups'
import {
  handlePresentationProduction,
  presentationProductionSummary,
} from './presentation-production'
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
  type PresentationInlineAsset,
} from '@wiswork/pptx-engine/presentation'
import { compilePresentationDeck } from '@wiswork/pptx-engine/presentation-compiler'

const MAX_RESPONSE_BYTES = 15 * 1024 * 1024
const locks = new Map<string, Promise<void>>()
const errorCodes = new Set([
  'busy',
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
  'evidence_source_unsupported',
  'evidence_changed',
  'attachment_conflict',
  'quota_exceeded',
  'digest_mismatch',
  'parse_failed',
  'page_not_ready',
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
  normalizeImage?: (
    bytes: Uint8Array,
  ) => Promise<{ bytes: Uint8Array; width: number; height: number }>
}): (body: unknown, signal: AbortSignal) => Promise<Uint8Array> {
  const pageBackups = createPresentationPageBackupService(options)
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
          'page_backup_begin',
          'page_backup_chunk',
          'page_backup_finish',
          'page_backup_status',
          'page_backup_read',
        ].includes(request.operation as string)
      )
        return boundedResponse(await pageBackups(request, signal))
      if (
        [
          'attachment_begin',
          'attachment_chunk',
          'attachment_finish',
          'attachment_list',
          'attachment_read',
          'attachment_asset',
          'attachment_list_assets',
        ].includes(request.operation as string)
      )
        return boundedResponse(await attachments(request, signal))
      if (
        ![
          ...presentationJobOperations,
          'compile',
          'get',
          'status',
          'resume',
          'save_plan',
          'get_plan',
          'production_begin',
          'production_rebuild_page',
          'production_status',
          'production_run',
          'production_page',
          'production_content_check',
          'production_page_reviews',
          'production_claim_evidence',
          'production_record_claim_review',
          'production_read_claim_review',
        ].includes(request.operation as string)
      )
        throw new Error('invalid_request')
      const allowedKeys = presentationJobOperations.includes(request.operation as string)
        ? ['operation', 'documentId', 'projectId', 'requestId']
        : request.operation === 'production_read_claim_review'
          ? ['operation', 'documentId', 'projectId', 'requestId', 'reviewId']
          : ['production_claim_evidence', 'production_record_claim_review'].includes(
                request.operation as string,
              )
            ? [
                'operation',
                'documentId',
                'projectId',
                'requestId',
                'pageId',
                'claimId',
                'sourceId',
                'offset',
                'maxChars',
                ...(request.operation === 'production_record_claim_review'
                  ? ['reviewId', 'evidenceDigest', 'outcome', 'notes']
                  : []),
              ]
            : request.operation === 'production_rebuild_page'
              ? [
                  'operation',
                  'documentId',
                  'projectId',
                  'parentRequestId',
                  'requestId',
                  'pageId',
                  'slide',
                ]
              : request.operation === 'production_begin'
                ? ['operation', 'documentId', 'projectId', 'requestId', 'planRevision', 'deck']
                : request.operation === 'production_status'
                  ? ['operation', 'documentId', 'projectId', 'requestId']
                  : request.operation === 'production_run'
                    ? ['operation', 'documentId', 'projectId', 'requestId']
                    : [
                          'production_page',
                          'production_content_check',
                          'production_page_reviews',
                        ].includes(request.operation as string)
                      ? ['operation', 'documentId', 'projectId', 'requestId', 'pageId']
                      : request.operation === 'compile'
                        ? [
                            'operation',
                            'documentId',
                            'projectId',
                            'requestId',
                            'deck',
                            'planRevision',
                          ]
                        : request.operation === 'resume'
                          ? ['operation', 'documentId', 'projectId', 'requestId']
                          : request.operation === 'save_plan'
                            ? ['operation', 'documentId', 'projectId', 'expectedRevision', 'plan']
                            : ['operation', 'documentId', 'projectId']
      const requiredKeys = allowedKeys.filter(
        (key) =>
          !(request.operation === 'compile' && ['projectId', 'planRevision'].includes(key)) &&
          !(request.operation === 'production_status' && key === 'requestId'),
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
      if (request.operation === 'compile' || request.operation === 'production_begin') {
        try {
          deck = parsePresentationDeck(request.deck)
        } catch {
          throw new Error('invalid_deck')
        }
        assertPresentationId(request.requestId)
      }
      if (
        [
          ...presentationJobOperations,
          'resume',
          'production_run',
          'production_page',
          'production_rebuild_page',
          'production_content_check',
          'production_page_reviews',
          'production_claim_evidence',
          'production_record_claim_review',
          'production_read_claim_review',
        ].includes(request.operation as string) ||
        (request.operation === 'production_status' && request.requestId !== undefined)
      )
        assertPresentationId(request.requestId)
      if (
        [
          'production_page',
          'production_rebuild_page',
          'production_content_check',
          'production_page_reviews',
          'production_claim_evidence',
          'production_record_claim_review',
        ].includes(request.operation as string)
      )
        assertPresentationId(request.pageId)
      if (
        ['production_claim_evidence', 'production_record_claim_review'].includes(
          request.operation as string,
        )
      ) {
        assertPresentationId(request.claimId)
        assertPresentationId(request.sourceId)
        if (
          !Number.isSafeInteger(request.offset) ||
          Number(request.offset) < 0 ||
          Number(request.offset) > 1000000 ||
          !Number.isSafeInteger(request.maxChars) ||
          Number(request.maxChars) < 1 ||
          Number(request.maxChars) > 8000
        )
          throw new Error('invalid_request')
      }
      if (
        ['production_record_claim_review', 'production_read_claim_review'].includes(
          request.operation as string,
        )
      )
        assertPresentationId(request.reviewId)
      if (request.operation === 'production_record_claim_review') {
        try {
          parsePresentationClaimReview({
            version: 1,
            projectId: request.projectId,
            requestId: request.requestId,
            reviewId: request.reviewId,
            pageId: request.pageId,
            claimId: request.claimId,
            sourceId: request.sourceId,
            offset: request.offset,
            maxChars: request.maxChars,
            evidenceDigest: request.evidenceDigest,
            outcome: request.outcome,
            notes: request.notes,
            reviewer: 'agent',
            planRevision: 1,
            inputDigest: '0'.repeat(64),
            planDigest: '0'.repeat(64),
            attachmentId: '0'.repeat(64),
            createdAt: '2026-01-01T00:00:00.000Z',
            checks: {
              support: 'agent_reviewed',
              sourceAuthority: 'not_verified',
              timeliness: 'not_verified',
              host: 'not_checked',
            },
          })
        } catch {
          throw new Error('invalid_request')
        }
      }
      if (request.operation === 'production_rebuild_page') {
        assertPresentationId(request.parentRequestId)
        if (request.parentRequestId === request.requestId) throw new Error('invalid_request')
      }
      if (
        ['compile', 'production_begin'].includes(request.operation as string) &&
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
        if (presentationJobOperations.includes(request.operation as string)) {
          const response = encode(
            handlePresentationJob(key, request, { store, compile, attachments }),
          )
          if (response.byteLength > 256 * 1024) throw new Error('output_too_large')
          return response
        }
        if (request.operation === 'production_run' && hasPresentationWorker(key))
          throw new Error('busy')
        if ((request.operation as string).startsWith('production_'))
          return boundedResponse(
            await handlePresentationProduction(request, { store, compile, attachments }, signal),
          )
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
          const productionRecord = store.production(
            projectId,
            documentId,
            activePresentationRequest(key),
          )
          const productionTasks = store
            .productionHistory(projectId, documentId)
            .slice(0, 32)
            .map((record) => {
              const summary = presentationProductionSummary(record)
              const job = store.productionJob(projectId, documentId, record.requestId)
              return {
                requestId: record.requestId,
                sequence: record.sequence,
                planRevision: record.plan.revision,
                status: summary.status,
                compiledCount: summary.compiledCount,
                total: summary.total,
                ...(job ? { jobState: job.state } : {}),
              }
            })
          const production = productionRecord
            ? presentationProductionSummary(productionRecord)
            : undefined
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
              ...(production ? { production, productionTasks } : {}),
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
            ...(production ? { production, productionTasks } : {}),
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
        // Keep compact references in the durable receipt. Resolve only against this document.
        const assets = []
        let imageBytes = 0
        for (const asset of inputDeck.assets) {
          checkAbort(signal)
          const resolved =
            'attachmentId' in asset
              ? {
                  ...((await attachments(
                    { operation: 'attachment_asset', documentId, attachmentId: asset.attachmentId },
                    signal,
                  )) as PresentationInlineAsset),
                  id: asset.id,
                }
              : asset
          if (!('base64' in resolved) || typeof resolved.base64 !== 'string')
            throw new Error('invalid_state')
          imageBytes += Buffer.byteLength(resolved.base64, 'base64')
          if (imageBytes > 8 * 1024 * 1024) throw new Error('output_too_large')
          assets.push(resolved)
        }
        checkAbort(signal)
        const compiled = await compile({ ...inputDeck, assets })
        checkAbort(signal)
        if (compiled.bytes.byteLength > 10 * 1024 * 1024) throw new Error('output_too_large')
        const sourceSlideIds = compiled.sourceSlideIds
        if (
          sourceSlideIds !== undefined &&
          (!Array.isArray(sourceSlideIds) ||
            sourceSlideIds.length !== inputDeck.slides.length ||
            new Set(sourceSlideIds).size !== sourceSlideIds.length ||
            sourceSlideIds.some(
              (id) =>
                typeof id !== 'string' ||
                !/^[1-9]\d*#$/.test(id) ||
                !Number.isSafeInteger(Number(id.slice(0, -1))) ||
                Number(id.slice(0, -1)) < 256 ||
                Number(id.slice(0, -1)) > 0xffffffff,
            ))
        )
          throw new Error('compile_failed')
        const result = {
          ...(sourceSlideIds
            ? {
                pages: inputDeck.slides.map((slide, i) => ({
                  id: slide.id,
                  title: slide.title,
                  sourceSlideId: sourceSlideIds[i]!,
                })),
              }
            : {}),
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
