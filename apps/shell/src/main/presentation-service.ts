import { handlePresentationDeliveryReport } from './presentation-delivery-report'
import { parsePresentationIssueActionInput } from '@wiswork/project-store/presentation-issue'
import {
  activePresentationRequest,
  handlePresentationJob,
  hasPresentationWorker,
  presentationJobOperations,
} from './presentation-jobs'
import { parsePresentationClaimReview } from '@wiswork/pptx-engine/presentation-claim-review'
import { createPresentationPageBackupService } from './presentation-page-backups'
import { createPresentationExistingPageBackupService } from './presentation-existing-page-backups'
import {
  assertCitedPresentationSourcesReady,
  handlePresentationProduction,
  presentationProductionSummary,
} from './presentation-production'
import { createPresentationAttachmentService } from './presentation-attachments'
import { auditPresentationSources } from './presentation-source-audit'
import {
  parsePresentationPlan,
  assertDeckMatchesPresentationPlan,
  assertBrandKitRevision,
} from '@wiswork/pptx-engine/presentation-plan'
import { resolve } from 'node:path'
import { PresentationStore, assertPresentationId } from '@wiswork/project-store'
import {
  parsePresentationDeck,
  type PresentationCompileReport,
  type PresentationInlineAsset,
} from '@wiswork/pptx-engine/presentation'
import { compilePresentationDeck } from '@wiswork/pptx-engine/presentation-compiler'
import { assertBrandLogoAsset, PresentationBrandLibrary } from './presentation-brand'
import { PresentationPreferenceLibrary } from './presentation-preferences'
import { PresentationCommentLibrary } from './presentation-comments'
import { convertPresentationToPdf } from './presentation-page-render'
import { PDFDocument } from 'pdf-lib'

const MAX_RESPONSE_BYTES = 15 * 1024 * 1024
const locks = new Map<string, Promise<void>>()
const errorCodes = new Set([
  'busy',
  'issue_changed',
  'invalid_request',
  'invalid_plan',
  'plan_mismatch',
  'revision_conflict',
  'invalid_deck',
  'invalid_state',
  'asset_unavailable',
  'source_unavailable',
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
  'animated_image_unsupported',
  'invalid_brand_kit',
  'remote_image_unavailable',
  'remote_image_source_conflict',
  'attachment_in_use',
  'page_not_ready',
  'renderer_unavailable',
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
  normalizeFirstFrame?: (
    bytes: Uint8Array,
  ) => Promise<{ bytes: Uint8Array; width: number; height: number }>
  renderPage?: (pptx: Uint8Array, signal: AbortSignal) => Promise<Uint8Array>
  renderPdf?: (pptx: Uint8Array, signal: AbortSignal) => Promise<Uint8Array>
}): (body: unknown, signal: AbortSignal) => Promise<Uint8Array> {
  const pageBackups = createPresentationPageBackupService(options)
  const existingPageBackups = createPresentationExistingPageBackupService(options)
  const attachments = createPresentationAttachmentService(options)
  const store = new PresentationStore(options.userDataPath)
  const brandLibrary = new PresentationBrandLibrary(options.userDataPath)
  const preferenceLibrary = new PresentationPreferenceLibrary(options.userDataPath)
  const commentLibrary = new PresentationCommentLibrary(options.userDataPath)
  const compile = options.compile ?? compilePresentationDeck
  const renderPdf = options.renderPdf ?? convertPresentationToPdf
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
      if (request.operation === 'export_pdf') {
        if (
          Object.keys(request).some(
            (key) => !['operation', 'documentId', 'projectId', 'requestId', 'source'].includes(key),
          ) ||
          ['documentId', 'projectId', 'requestId'].some((key) => !Object.hasOwn(request, key)) ||
          typeof request.documentId !== 'string' ||
          !request.documentId.trim() ||
          request.documentId.length > 2048 ||
          (request.source !== undefined &&
            request.source !== 'compiled' &&
            request.source !== 'production')
        )
          throw new Error('invalid_request')
        assertPresentationId(request.projectId)
        assertPresentationId(request.requestId)
        const source = request.source ?? 'compiled'
        let pdf: Buffer
        let slideCount: number
        if (source === 'production') {
          const production = store.production(
            request.projectId as string,
            request.documentId,
            request.requestId as string,
          )
          if (!production) throw new Error('not_found')
          if (production.pages.some((page) => page.state !== 'compiled' || !page.result))
            throw new Error('page_not_ready')
          const merged = await PDFDocument.create()
          let inputBytes = 0
          for (const page of production.pages) {
            checkAbort(signal)
            const pptx = Buffer.from(page.result!.pptxBase64, 'base64')
            inputBytes += pptx.length
            if (inputBytes > 10 * 1024 * 1024) throw new Error('output_too_large')
            const rendered = await renderPdf(pptx, signal)
            checkAbort(signal)
            let onePage: PDFDocument
            try {
              onePage = await PDFDocument.load(rendered)
            } catch {
              throw new Error('renderer_unavailable')
            }
            if (onePage.getPageCount() !== 1) throw new Error('renderer_unavailable')
            const [copied] = await merged.copyPages(onePage, [0])
            merged.addPage(copied)
          }
          slideCount = production.pages.length
          pdf = Buffer.from(await merged.save())
        } else {
          const record = store.request(
            request.projectId as string,
            request.documentId,
            request.requestId as string,
          )
          if (record?.status !== 'compiled') throw new Error('not_found')
          const compiled = record.result as {
            pptxBase64: string
            report: PresentationCompileReport
          }
          pdf = Buffer.from(await renderPdf(Buffer.from(compiled.pptxBase64, 'base64'), signal))
          slideCount = compiled.report.slideCount
        }
        checkAbort(signal)
        if (
          pdf.length < 16 ||
          pdf.length > 10 * 1024 * 1024 ||
          !pdf.subarray(0, 5).equals(Buffer.from('%PDF-')) ||
          !pdf.subarray(Math.max(0, pdf.length - 1024)).includes(Buffer.from('%%EOF'))
        )
          throw new Error('renderer_unavailable')
        let pageCount: number
        try {
          pageCount = (await PDFDocument.load(pdf)).getPageCount()
        } catch {
          throw new Error('renderer_unavailable')
        }
        if (pageCount !== slideCount) throw new Error('renderer_unavailable')
        return boundedResponse({
          status: 'exported',
          source,
          projectId: request.projectId,
          requestId: request.requestId,
          slideCount: pageCount,
          pdfBase64: pdf.toString('base64'),
        })
      }
      if (
        ['comment_list', 'comment_add', 'comment_resolve'].includes(request.operation as string)
      ) {
        const required =
          request.operation === 'comment_add'
            ? [
                'operation',
                'documentId',
                'projectId',
                'expectedRevision',
                'planRevision',
                'comment',
              ]
            : request.operation === 'comment_resolve'
              ? ['operation', 'documentId', 'projectId', 'expectedRevision', 'commentId']
              : ['operation', 'documentId', 'projectId']
        if (
          Object.keys(request).sort().join(',') !== required.sort().join(',') ||
          typeof request.documentId !== 'string' ||
          !request.documentId ||
          request.documentId.length > 2048
        )
          throw new Error('invalid_request')
        assertPresentationId(request.projectId)
        const record = store.plan(request.projectId, request.documentId)
        if (!record) throw new Error('not_found')
        if (request.operation === 'comment_list')
          return boundedResponse(commentLibrary.list(request.documentId, request.projectId))
        if (request.operation === 'comment_resolve')
          return boundedResponse(
            commentLibrary.resolve(
              request.documentId,
              request.projectId,
              request.expectedRevision as number,
              request.commentId as string,
            ),
          )
        return boundedResponse(
          commentLibrary.add(
            request.documentId,
            request.projectId,
            request.expectedRevision as number,
            request.planRevision as number,
            request.comment,
            { revision: record.revision, plan: parsePresentationPlan(record.plan) },
          ),
        )
      }
      if (
        request.operation === 'preference_save' ||
        request.operation === 'preference_list' ||
        request.operation === 'preference_delete'
      ) {
        const required =
          request.operation === 'preference_save'
            ? ['operation', 'documentId', 'preference']
            : request.operation === 'preference_delete'
              ? ['operation', 'documentId', 'projectId', 'changeId']
              : ['operation', 'documentId', 'projectId']
        if (
          Object.keys(request).sort().join(',') !== required.sort().join(',') ||
          typeof request.documentId !== 'string' ||
          !request.documentId ||
          request.documentId.length > 2048
        )
          throw new Error('invalid_request')
        if (request.operation === 'preference_save')
          return boundedResponse({
            preference: preferenceLibrary.save(request.documentId, request.preference),
          })
        if (request.operation === 'preference_delete')
          return boundedResponse({
            deleted: preferenceLibrary.delete(
              request.documentId,
              request.projectId as string,
              request.changeId as string,
            ),
          })
        return boundedResponse({
          preferences: preferenceLibrary.list(request.documentId, request.projectId as string),
        })
      }
      if (
        ['brand_kit_save', 'brand_kit_get', 'brand_kit_list'].includes(request.operation as string)
      ) {
        const operation = request.operation
        const required =
          operation === 'brand_kit_save'
            ? ['operation', 'documentId', 'expectedRevision', 'brandKit']
            : operation === 'brand_kit_get'
              ? ['operation', 'documentId', 'brandKitId', 'revision']
              : ['operation', 'documentId']
        if (
          Object.keys(request).sort().join(',') !== required.sort().join(',') ||
          typeof request.documentId !== 'string' ||
          !request.documentId ||
          request.documentId.length > 2048
        )
          throw new Error('invalid_request')
        if (operation === 'brand_kit_list')
          return boundedResponse({ brandKits: brandLibrary.list() })
        if (operation === 'brand_kit_get') {
          const kit = brandLibrary.get(request.brandKitId as string, request.revision as number)
          if (!kit) throw new Error('not_found')
          return boundedResponse({ brandKit: kit })
        }
        try {
          return boundedResponse({
            brandKit: brandLibrary.save(request.expectedRevision as number, request.brandKit),
          })
        } catch (error) {
          if (error instanceof Error && error.message === 'presentation_brand_kit_invalid')
            throw new Error('invalid_brand_kit', { cause: error })
          throw error
        }
      }
      if (
        [
          'existing_page_backup_begin',
          'existing_page_backup_chunk',
          'existing_page_backup_finish',
          'existing_page_backup_status',
          'existing_page_backup_read',
          'existing_page_backup_render',
          'existing_page_backup_release',
          'existing_page_backup_abandon',
          'existing_page_backup_list',
        ].includes(request.operation as string)
      )
        return boundedResponse(await existingPageBackups(request, signal))
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
          'attachment_extract_first_frame',
          'attachment_delete',
          'attachment_attest_license',
          'attachment_revoke_license',
          'attachment_import_url',
          'attachment_list',
          'attachment_read',
          'attachment_asset',
          'attachment_original',
          'attachment_list_assets',
          'attachment_metadata',
          'attachment_match_excerpt',
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
          'audit_sources',
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
          'production_delivery_report',
          'production_record_issue_action',
        ].includes(request.operation as string)
      )
        throw new Error('invalid_request')
      const allowedKeys = presentationJobOperations.includes(request.operation as string)
        ? ['operation', 'documentId', 'projectId', 'requestId']
        : ['production_delivery_report', 'production_record_issue_action'].includes(
              request.operation as string,
            )
          ? [
              'operation',
              'documentId',
              'projectId',
              'requestId',
              ...(request.operation === 'production_record_issue_action'
                ? ['expectedRevision', 'action']
                : []),
            ]
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
          'production_delivery_report',
          'production_record_issue_action',
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
      if (request.operation === 'production_record_issue_action') {
        parsePresentationIssueActionInput(request.action)
        if (!Number.isSafeInteger(request.expectedRevision) || Number(request.expectedRevision) < 0)
          throw new Error('invalid_request')
      }
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
        if (
          ['production_delivery_report', 'production_record_issue_action'].includes(
            request.operation as string,
          )
        )
          return boundedResponse(
            await handlePresentationDeliveryReport(request, store, attachments, signal),
          )
        if ((request.operation as string).startsWith('production_'))
          return boundedResponse(
            await handlePresentationProduction(request, { store, compile, attachments }, signal),
          )
        if (request.operation === 'save_plan' || request.operation === 'get_plan') {
          if (request.operation === 'save_plan') {
            const previousPlan = store.plan(projectId, documentId)
            if (previousPlan && previousPlan.revision === request.expectedRevision) {
              try {
                assertBrandKitRevision(parsePresentationPlan(previousPlan.plan), plan!)
              } catch {
                throw new Error('invalid_plan')
              }
            }
          }
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
        if (request.operation === 'audit_sources') {
          const record = store.plan(projectId, documentId)
          if (!record) throw new Error('not_found')
          const plan = parsePresentationPlan(record.plan)
          const sources = await auditPresentationSources(plan, documentId, attachments, signal)
          return boundedResponse({
            projectId,
            planRevision: record.revision,
            sources,
            checks: {
              support: 'not_verified',
              sourceAuthority: 'not_verified',
              timeliness: 'not_verified',
            },
          })
        }
        if (request.operation === 'status') {
          let reviewComments:
            | {
                revision: number
                openCount: number
                resolvedCount: number
                recent: {
                  id: string
                  targetKind: string
                  targetId: string
                  authorLabel: string
                  text: string
                  state: string
                  planRevision: number
                  createdAt: string
                }[]
              }
            | undefined
          let commentsUnavailable = false
          try {
            const ledger = commentLibrary.list(documentId, projectId)
            reviewComments = {
              revision: ledger.revision,
              openCount: ledger.comments.filter((comment) => comment.state === 'open').length,
              resolvedCount: ledger.comments.filter((comment) => comment.state === 'resolved')
                .length,
              recent: ledger.comments.slice(-8).map((comment) => ({
                id: comment.id,
                targetKind: comment.targetKind,
                targetId: comment.targetId,
                authorLabel: comment.authorLabel,
                text: comment.text.slice(0, 400),
                state: comment.state,
                planRevision: comment.planRevision,
                createdAt: comment.createdAt,
              })),
            }
          } catch {
            commentsUnavailable = true
          }
          const commentStatus = commentsUnavailable
            ? { commentsUnavailable: true }
            : reviewComments && reviewComments.openCount + reviewComments.resolvedCount > 0
              ? { reviewComments }
              : {}
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
              const lastEvent = job?.events.at(-1)
              return {
                requestId: record.requestId,
                sequence: record.sequence,
                planRevision: record.plan.revision,
                status: summary.status,
                compiledCount: summary.compiledCount,
                total: summary.total,
                ...(job ? { jobState: job.state } : {}),
                ...(lastEvent
                  ? {
                      lastEvent: {
                        type: lastEvent.type,
                        createdAt: lastEvent.createdAt,
                        ...('pageId' in lastEvent ? { pageId: lastEvent.pageId } : {}),
                      },
                    }
                  : {}),
              }
            })
          const production = productionRecord
            ? presentationProductionSummary(productionRecord)
            : undefined
          const savedPlan = store.plan(projectId, documentId)
          const plan = savedPlan
            ? {
                revision: savedPlan.revision,
                value: parsePresentationPlan(savedPlan.plan),
                ...(savedPlan.revisions ? { revisions: savedPlan.revisions } : {}),
              }
            : undefined
          let sourcePreparation:
            | {
                sourceId: string
                attachmentId: string
                status:
                  | 'excerpt_matched'
                  | 'uploading'
                  | 'failed'
                  | 'missing'
                  | 'unsupported'
                  | 'excerpt_mismatch'
                  | 'excerpt_missing'
              }[]
            | undefined
          let sourcePreparationUnavailable = false
          if (plan) {
            const references = plan.value.sources.flatMap((source) => {
              const match = /^attachment:([a-f0-9]{64})$/.exec(source.uri)
              return match
                ? [{ sourceId: source.id, attachmentId: match[1]!, excerpt: source.excerpt }]
                : []
            })
            if (references.length) {
              try {
                sourcePreparation = []
                for (const reference of references) {
                  try {
                    const item = (await attachments(
                      {
                        operation: 'attachment_metadata',
                        documentId,
                        attachmentId: reference.attachmentId,
                      },
                      signal,
                    )) as {
                      attachmentId: string
                      status: 'ready' | 'uploading' | 'failed'
                      kind?: 'text' | 'image'
                    }
                    if (item.attachmentId !== reference.attachmentId)
                      throw new Error('invalid_state')
                    let status: NonNullable<typeof sourcePreparation>[number]['status'] =
                      item.status === 'ready' ? 'unsupported' : item.status
                    if (item.status === 'ready' && item.kind === 'text') {
                      const match = (await attachments(
                        {
                          operation: 'attachment_match_excerpt',
                          documentId,
                          attachmentId: reference.attachmentId,
                          excerpt: reference.excerpt,
                        },
                        signal,
                      )) as { attachmentId?: unknown; status?: unknown }
                      if (match.attachmentId !== reference.attachmentId)
                        throw new Error('invalid_state')
                      status = {
                        found: 'excerpt_matched',
                        not_found: 'excerpt_mismatch',
                        empty_excerpt: 'excerpt_missing',
                        not_ready: 'uploading',
                        unsupported: 'unsupported',
                      }[String(match.status)] as typeof status
                      if (!status) throw new Error('invalid_state')
                    }
                    sourcePreparation.push({
                      sourceId: reference.sourceId,
                      attachmentId: reference.attachmentId,
                      status,
                    })
                  } catch (error) {
                    if (!(error instanceof Error) || error.message !== 'not_found') throw error
                    sourcePreparation.push({
                      sourceId: reference.sourceId,
                      attachmentId: reference.attachmentId,
                      status: 'missing',
                    })
                  }
                }
              } catch {
                sourcePreparation = undefined
                sourcePreparationUnavailable = true
              }
            }
          }
          const preparationStatus = sourcePreparationUnavailable
            ? { sourcePreparationUnavailable: true }
            : sourcePreparation
              ? { sourcePreparation }
              : {}
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
              ...commentStatus,
              ...preparationStatus,
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
            ...commentStatus,
            ...preparationStatus,
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
          let plan: ReturnType<typeof parsePresentationPlan>
          try {
            plan = parsePresentationPlan(record.plan.plan)
            assertDeckMatchesPresentationPlan(inputDeck, plan)
          } catch {
            throw new Error('plan_mismatch')
          }
          const sourceReadiness = new Map<string, Promise<boolean>>()
          for (const slide of inputDeck.slides)
            await assertCitedPresentationSourcesReady(
              plan,
              slide,
              documentId,
              attachments,
              signal,
              sourceReadiness,
            )
        }
        // Keep compact references in the durable receipt. Resolve only against this document.
        const assets = []
        let imageBytes = 0
        for (const asset of inputDeck.assets) {
          checkAbort(signal)
          let resolved = asset
          if ('attachmentId' in asset) {
            try {
              resolved = {
                ...((await attachments(
                  { operation: 'attachment_asset', documentId, attachmentId: asset.attachmentId },
                  signal,
                )) as PresentationInlineAsset),
                id: asset.id,
              }
            } catch (error) {
              checkAbort(signal)
              if (
                error instanceof Error &&
                ['invalid_state', 'digest_mismatch'].includes(error.message)
              )
                throw new Error('asset_unavailable', { cause: error })
              throw error
            }
          }
          if (!('base64' in resolved) || typeof resolved.base64 !== 'string')
            throw new Error('invalid_state')
          imageBytes += Buffer.byteLength(resolved.base64, 'base64')
          if (imageBytes > 8 * 1024 * 1024) throw new Error('output_too_large')
          assets.push(resolved)
        }
        checkAbort(signal)
        if (record.plan) assertBrandLogoAsset(parsePresentationPlan(record.plan.plan), assets)
        // Only attachment_asset can add evidence after document-bound digest validation.
        const compiled = await compile({ ...inputDeck, assets }, { trustedAssetEvidence: true })
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
