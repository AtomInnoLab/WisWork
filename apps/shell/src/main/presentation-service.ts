import { createHash } from 'node:crypto'
import { canonicalPresentationValue } from '@wiswork/project-store/presentation-canonical'
import {
  parseSavedPresentationPreference,
  parsePresentationPreferenceSource,
  parsePresentationPreferenceOrigin,
} from '@wiswork/pptx-engine/presentation-preference'
import { parsePresentationManualObservationShape } from '@wiswork/pptx-engine/presentation-manual-observation'
import {
  registerPresentationProjectWork,
  type PresentationProjectWork,
} from './presentation-project-work'
import {
  capturePresentationProjectWriteLease,
  capturePresentationProjectReadLease,
  capturePresentationProjectCreationLease,
  capturePresentationProjectAsyncReadLease,
} from './presentation-project-write-lease'
import {
  createPresentationProjectLifecycleService,
  presentationProjectLifecycleOperations,
} from './presentation-project-lifecycle'
import { PresentationManualObservationLibrary } from './presentation-manual-observations'
import { buildPresentationFeedbackComparison } from '@wiswork/pptx-engine/presentation-feedback-comparison'
import { parsePresentationProductionFeedbackPages } from '@wiswork/project-store/presentation-feedback'
import { createPresentationTeamService } from './presentation-team'
import type { PresentationTeamContext } from '@wiswork/pptx-engine/presentation-team'
import { parsePresentationSourceAssessment } from '@wiswork/project-store/presentation-source-assessment'
import { PresentationResearchStore } from '@wiswork/project-store/presentation-research-store'
import { readBoundPresentationResearch } from './presentation-research-plan-binding'
import { createPresentationResearchService } from './presentation-research'
import { createPresentationDeliveryBundleService } from './presentation-delivery-bundles'
import { MAX_PRESENTATION_SOURCE_TEXT_CHARS } from '@wiswork/pptx-engine/presentation-source-limits'
import { readPresentationImportSource } from './presentation-import-source'
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
import { createPresentationPackageBackupService } from './presentation-package-backups'
import { createPresentationMasterBackupService } from './presentation-master-backups'
import {
  assertCitedPresentationSourcesReady,
  handlePresentationProduction,
  presentationProductionSummary,
} from './presentation-production'
import { createPresentationAttachmentService } from './presentation-attachments'
import { canonicalSourceLocator, matchesFetchedSourceUrl } from './presentation-source-audit'
import {
  parsePresentationPlan,
  assertDeckMatchesPresentationPlan,
  assertBrandKitRevision,
  presentationSourceAttachmentId,
} from '@wiswork/pptx-engine/presentation-plan'
import { resolve } from 'node:path'
import {
  PresentationLifecycleStore,
  PresentationStore,
  assertPresentationId,
} from '@wiswork/project-store'
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
import { isInstalledFontFamily } from '@wiswork/font-metrics'
import { handlePresentationSourceAudit } from './presentation-source-audit-history'
import { presentationSourceAuditHistory } from '@wiswork/project-store/presentation-source-audit'

const readonlyProductionOperations = new Set([
  'production_status',
  'production_page',
  'production_content_check',
  'production_page_reviews',
  'production_claim_evidence',
  'production_read_claim_review',
  'production_delivery_report',
  'production_feedback_read',
  'production_feedback_compare',
])
const synchronousLibraryOperations = new Set([
  'comment_list',
  'comment_add',
  'comment_resolve',
  'manual_observation_begin',
  'manual_observation_complete',
  'manual_observation_get',
  'manual_observation_list',
  'manual_observation_delete',
  'preference_save_observation',
  'preference_get',
  'preference_import',
  'preference_save',
  'preference_list',
  'preference_delete',
])
const ordinaryProjectReadOperations = new Set([
  'get',
  'status',
  'get_plan',
  'read_import_source',
  'read_source_audit',
])
const ordinaryProjectWriteOperations = new Set([
  'compile',
  'resume',
  'save_plan',
  'accept_plan',
  'set_plan_page_lock',
  'audit_sources',
])
const MAX_RESPONSE_BYTES = 15 * 1024 * 1024
const locks = new Map<string, Promise<void>>()
async function acquireProjectLock(
  root: string,
  projectId: string,
  signal?: AbortSignal,
): Promise<() => void> {
  const key = `${resolve(root)}\0${projectId}`
  const previous = locks.get(key) ?? Promise.resolve()
  let release!: () => void
  const current = new Promise<void>((done) => {
    release = done
  })
  locks.set(key, current)
  let acquired = true
  if (signal) {
    acquired = signal.aborted
      ? false
      : await new Promise<boolean>((resolve) => {
          const onAbort = () => resolve(false)
          signal.addEventListener('abort', onAbort, { once: true })
          void previous.then(() => {
            signal.removeEventListener('abort', onAbort)
            resolve(true)
          })
        })
  } else await previous
  if (!acquired || signal?.aborted) {
    // A cancelled waiter remains in the queue until the active owner exits.
    void previous.then(() => {
      release()
      if (locks.get(key) === current) locks.delete(key)
    })
    throw new Error('aborted')
  }
  return () => {
    release()
    if (locks.get(key) === current) locks.delete(key)
  }
}
const errorCodes = new Set([
  'presentation_master_backup_invalid',
  'presentation_master_backup_capacity',
  'presentation_package_backup_invalid',
  'presentation_package_backup_capacity',
  'project_not_found',
  'project_deleting',
  'project_deleted',
  'access_denied',
  'busy',
  'issue_changed',
  'invalid_request',
  'invalid_plan',
  'plan_mismatch',
  'revision_conflict',
  'research_binding_invalid',
  'research_unavailable',
  'upgrade_required',
  'record_deleted',
  'record_running',
  'record_not_running',
  'record_protected',
  'cleanup_quota_exceeded',
  'acceptance_capacity',
  'plan_revision_unavailable',
  'page_locked',
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
  'evidence_source_mismatch',
  'evidence_changed',
  'evidence_excerpt_not_found',
  'evidence_locator_mismatch',
  'evidence_image_backed_unverified',
  'attachment_conflict',
  'revision_conflict',
  'quota_exceeded',
  'digest_mismatch',
  'parse_failed',
  'animated_image_unsupported',
  'invalid_brand_kit',
  'remote_image_unavailable',
  'remote_image_source_conflict',
  'remote_webpage_unavailable',
  'remote_webpage_source_conflict',
  'font_unavailable',
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
  fetchImage?: (url: string, signal: AbortSignal) => Promise<Response | null>
  fetchPage?: (url: string, signal: AbortSignal) => Promise<Response | null>
}): (body: unknown, signal: AbortSignal, context?: PresentationTeamContext) => Promise<Uint8Array> {
  const lifecycle = createPresentationProjectLifecycleService(options)
  const existingPageBackups = createPresentationExistingPageBackupService(options)
  const packageBackups = createPresentationPackageBackupService(options)
  const masterBackups = createPresentationMasterBackupService(options)
  const attachments = createPresentationAttachmentService(options)
  const store = new PresentationStore(options.userDataPath)
  const lifecycleStore = new PresentationLifecycleStore(options.userDataPath)
  const captureFactoryLease = (
    scope: Readonly<{ documentId: string; projectId: string }>,
    mode: 'read' | 'write',
    signal: AbortSignal,
  ) => {
    const input = {
      store: lifecycleStore,
      scope,
      signal,
      readExistingProject: () => store.projectScope(scope.projectId, scope.documentId),
    }
    if (mode === 'read') return capturePresentationProjectReadLease(input)
    const lease = capturePresentationProjectWriteLease(input)
    return { assertCurrent: lease.assertWritable }
  }
  const pageBackups = createPresentationPageBackupService({
    ...options,
    captureProjectLease: ({ scope, operation, signal }) => {
      const input = {
        store: lifecycleStore,
        scope,
        signal,
        readExistingProject: () => store.projectScope(scope.projectId, scope.documentId),
      }
      return ['page_backup_status', 'page_backup_read'].includes(operation)
        ? capturePresentationProjectReadLease(input)
        : capturePresentationProjectWriteLease(input)
    },
  })
  const team = createPresentationTeamService({
    userDataPath: options.userDataPath,
    captureProjectLease: captureFactoryLease,
    readPlan: (documentId, projectId) => {
      const record = store.plan(projectId, documentId)
      return record
        ? { revision: record.revision, plan: parsePresentationPlan(record.plan) }
        : undefined
    },
    acquireProjectLock: (projectId, signal) =>
      acquireProjectLock(options.userDataPath, projectId, signal),
  })
  const deliveryBundles = createPresentationDeliveryBundleService({
    ...options,
    captureProjectLease: captureFactoryLease,
    acquireProjectLock: (projectId, signal) =>
      acquireProjectLock(options.userDataPath, projectId, signal),
  })
  const research = createPresentationResearchService({
    userDataPath: options.userDataPath,
    attachments,
    captureProjectLease: async ({ scope, operation, signal }) => {
      if (operation === 'research_build')
        return capturePresentationProjectCreationLease({ store: lifecycleStore, scope, signal })
      const readOnly = [
        'research_list',
        'research_latest',
        'research_read',
        'research_delete_status',
      ].includes(operation)
      if (readOnly) {
        let initial: ReturnType<PresentationLifecycleStore['readControl']>
        try {
          initial = lifecycleStore.readControl(scope)
        } catch (error) {
          if (!(error instanceof Error) || error.message !== 'document_mismatch') throw error
          const summary = await researchStore.summary(scope.documentId, scope.projectId)
          checkAbort(signal)
          if (summary.totalRecords || (summary.version === 2 && summary.lastSequence > 0))
            throw error
          return Object.freeze({
            scope,
            revision: undefined,
            assertCurrent: () => checkAbort(signal),
          })
        }
        const assertCurrent = () => {
          checkAbort(signal)
          const current = lifecycleStore.readControl(scope)
          if (current?.revision !== initial?.revision) throw Error('revision_conflict')
          if (current && current.state !== 'active') throw Error('project_' + current.state)
        }
        assertCurrent()
        // The strict research envelope proves private ownership independently of production metadata.
        await researchStore.summary(scope.documentId, scope.projectId)
        assertCurrent()
        return Object.freeze({ scope, revision: initial?.revision, assertCurrent })
      }
      let empty = false
      let initialRevision: number | undefined
      try {
        initialRevision = lifecycleStore.read(scope)?.revision
      } catch (error) {
        if (!readOnly || !(error instanceof Error) || error.message !== 'document_mismatch')
          throw error
        const summary = await researchStore.summary(scope.documentId, scope.projectId)
        checkAbort(signal)
        if (summary.totalRecords || (summary.version === 2 && summary.lastSequence > 0)) throw error
        return Object.freeze({
          scope,
          revision: undefined,
          assertCurrent: () => checkAbort(signal),
        })
      }
      const assertInitial = () => {
        checkAbort(signal)
        if (initialRevision === undefined) {
          if (lifecycleStore.read(scope) !== undefined) throw Error('revision_conflict')
        } else lifecycleStore.assertActive(scope, initialRevision)
      }
      const readLease = await capturePresentationProjectAsyncReadLease({
        store: lifecycleStore,
        scope,
        signal,
        readExistingProject: async () => {
          if (!readOnly) {
            const project = store.projectScope(scope.projectId, scope.documentId)
            if (project) return project
          }
          const summary = await researchStore.summary(scope.documentId, scope.projectId)
          empty = summary.totalRecords === 0 && !(summary.version === 2 && summary.lastSequence > 0)
          return empty ? undefined : scope
        },
      }).catch((error) => {
        if (readOnly && empty && error instanceof Error && error.message === 'project_not_found') {
          assertInitial()
          return Object.freeze({ scope, revision: initialRevision, assertCurrent: assertInitial })
        }
        throw error
      })
      readLease.assertCurrent()
      if (readOnly) return readLease
      return capturePresentationProjectWriteLease({
        store: lifecycleStore,
        scope,
        signal,
        readExistingProject: () => scope,
      })
    },
    acquireProjectLock: (projectId, signal) =>
      acquireProjectLock(options.userDataPath, projectId, signal),
    assertRecordUnprotected: async (documentId, projectId, ledgerId, signal) => {
      const plans: unknown[] = []
      const current = store.plan(projectId, documentId, true)
      if (current) {
        plans.push(current.plan)
        for (const event of current.revisions ?? []) {
          const revision = store.planRevision(projectId, documentId, event.revision)
          if (!revision) throw new Error('invalid_state')
          plans.push(revision.plan)
        }
      }
      for (const record of store.productionHistory(projectId, documentId))
        plans.push(record.plan.plan)
      for (const receipt of store.history(projectId, documentId, true))
        if (receipt.plan) plans.push(receipt.plan.plan)
      for (const value of plans) {
        let plan: ReturnType<typeof parsePresentationPlan>
        try {
          plan = parsePresentationPlan(value)
        } catch {
          throw new Error('invalid_state')
        }
        if (plan.projectId !== projectId) throw new Error('invalid_state')
        if (plan.research?.ledgerId === ledgerId) throw new Error('record_protected')
      }
      await deliveryBundles.assertResearchCleanupAvailable(documentId, projectId)
      checkAbort(signal)
    },
  })
  const researchStore = new PresentationResearchStore(options.userDataPath)
  const brandLibrary = new PresentationBrandLibrary(options.userDataPath)
  const preferenceLibrary = new PresentationPreferenceLibrary(options.userDataPath)
  const manualObservations = new PresentationManualObservationLibrary(options.userDataPath)
  const commentLibrary = new PresentationCommentLibrary(options.userDataPath)
  const compile =
    options.compile ??
    ((deck, settings) =>
      compilePresentationDeck(deck, {
        ...settings,
        fontAvailable: isInstalledFontFamily,
      }))
  const renderPdf = options.renderPdf ?? convertPresentationToPdf
  return async (body, signal, context) => {
    let foregroundWork: PresentationProjectWork | undefined
    try {
      checkAbort(signal)
      if (
        !body ||
        typeof body !== 'object' ||
        Array.isArray(body) ||
        Buffer.byteLength(JSON.stringify(body)) > 256 * 1024
      )
        throw new Error('invalid_request')
      let request = body as Record<string, unknown>
      if (
        typeof request.operation === 'string' &&
        (request.operation.startsWith('production_') ||
          request.operation === 'export_pdf' ||
          ordinaryProjectReadOperations.has(request.operation) ||
          ordinaryProjectWriteOperations.has(request.operation) ||
          synchronousLibraryOperations.has(request.operation))
      )
        request = structuredClone(request)
      if (
        presentationProjectLifecycleOperations.some((operation) => operation === request.operation)
      )
        return boundedResponse(lifecycle(request))
      if (typeof request.operation === 'string' && request.operation.startsWith('team_'))
        return boundedResponse(await team(request, context, signal))
      if (typeof request.operation === 'string' && request.operation.startsWith('research_'))
        return boundedResponse(await research(request, signal))
      if (
        typeof request.operation === 'string' &&
        request.operation.startsWith('delivery_bundle_')
      ) {
        return boundedResponse(await deliveryBundles(request, signal))
      }
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
        const scope = { projectId: request.projectId as string, documentId: request.documentId }
        const existing = store.projectScope(scope.projectId, scope.documentId)
        const control = lifecycleStore.read(scope)
        if (control) lifecycleStore.assertActive(scope, control.revision)
        if (!existing && !control) throw new Error('not_found')
        foregroundWork = registerPresentationProjectWork({
          scope: { root: options.userDataPath, ...scope },
          signal,
        })
        signal = foregroundWork.signal
        const lease = capturePresentationProjectReadLease({
          store: lifecycleStore,
          scope,
          readExistingProject: (bound) => store.projectScope(bound.projectId, bound.documentId),
          signal,
        })
        const assertCurrent = lease.assertCurrent
        try {
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
            assertCurrent()
            let inputBytes = 0
            for (const page of production.pages) {
              checkAbort(signal)
              const pptx = Buffer.from(page.result!.pptxBase64, 'base64')
              inputBytes += pptx.length
              if (inputBytes > 10 * 1024 * 1024) throw new Error('output_too_large')
              const rendered = await renderPdf(pptx, signal)
              assertCurrent()
              checkAbort(signal)
              let onePage: PDFDocument
              try {
                onePage = await PDFDocument.load(rendered)
              } catch {
                throw new Error('renderer_unavailable')
              }
              assertCurrent()
              if (onePage.getPageCount() !== 1) throw new Error('renderer_unavailable')
              const [copied] = await merged.copyPages(onePage, [0])
              assertCurrent()
              merged.addPage(copied)
            }
            slideCount = production.pages.length
            pdf = Buffer.from(await merged.save())
            assertCurrent()
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
            assertCurrent()
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
          assertCurrent()
          if (pageCount !== slideCount) throw new Error('renderer_unavailable')
          assertCurrent()
          return boundedResponse({
            status: 'exported',
            source,
            projectId: request.projectId,
            requestId: request.requestId,
            slideCount: pageCount,
            pdfBase64: pdf.toString('base64'),
          })
        } catch (error) {
          assertCurrent()
          throw error
        }
      }
      const registerLibrary = (scope: { documentId: string; projectId: string }) => {
        foregroundWork = registerPresentationProjectWork({
          scope: { root: options.userDataPath, ...scope },
          signal,
        })
        signal = foregroundWork.signal
      }
      const scopedGuard =
        (scope: { documentId: string; projectId: string }, check: () => void) =>
        (actual: Readonly<{ documentId: string; projectId: string }>) => {
          if (actual.documentId !== scope.documentId || actual.projectId !== scope.projectId)
            throw new Error('invalid_request')
          check()
        }
      const readLibrary = <T>(
        scope: { documentId: string; projectId: string },
        load: () => T,
        proven: (value: T) => boolean,
      ) => {
        let value!: T,
          loaded = false,
          empty = false
        const read = () => {
          if (!loaded) {
            value = load()
            loaded = true
          }
          return value
        }
        let lease
        try {
          lease = capturePresentationProjectReadLease({
            store: lifecycleStore,
            scope,
            signal,
            readExistingProject: () => {
              const project = store.projectScope(scope.projectId, scope.documentId)
              const data = read()
              if (project || proven(data)) return scope
              empty = true
              return undefined
            },
          })
        } catch (error) {
          if (error instanceof Error && error.message === 'document_mismatch') {
            const data = read()
            if (!proven(data)) return { value: data, assertCurrent: () => checkAbort(signal) }
          }
          if (!empty) throw error
          // Empty queries prove no ownership and grant no write permission; keep control absence fixed.
          const assertCurrent = () => {
            checkAbort(signal)
            if (lifecycleStore.read(scope) !== undefined) throw new Error('revision_conflict')
          }
          assertCurrent()
          return { value, assertCurrent }
        }
        lease.assertCurrent()
        read()
        lease.assertCurrent()
        return { value, assertCurrent: lease.assertCurrent }
      }
      const libraryWrite = (
        scope: { documentId: string; projectId: string },
        proof: boolean,
        creation = false,
      ) => {
        const project = store.projectScope(scope.projectId, scope.documentId)
        const lease = creation
          ? capturePresentationProjectCreationLease({ store: lifecycleStore, scope, signal })
          : capturePresentationProjectWriteLease({
              store: lifecycleStore,
              scope,
              signal,
              readExistingProject: () => project || (proof ? scope : undefined),
            })
        return scopedGuard(scope, lease.assertWritable)
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
        const bound = { documentId: record.documentId, projectId: record.projectId }
        registerLibrary(bound)
        if (request.operation === 'comment_list') {
          const access = readLibrary(
            bound,
            () => commentLibrary.list(bound.documentId, bound.projectId),
            () => true,
          )
          access.assertCurrent()
          return boundedResponse(access.value)
        }
        const readLease = capturePresentationProjectReadLease({
          store: lifecycleStore,
          scope: bound,
          signal,
          readExistingProject: () => record,
        })
        let admittedGuard: ReturnType<typeof libraryWrite> | undefined
        const guard = scopedGuard(bound, () => {
          if (!admittedGuard) {
            readLease.assertCurrent()
            admittedGuard = libraryWrite(bound, true)
          }
          admittedGuard(bound)
        })
        if (request.operation === 'comment_resolve')
          return boundedResponse(
            commentLibrary.resolve(
              request.documentId,
              request.projectId,
              request.expectedRevision as number,
              request.commentId as string,
              guard,
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
            guard,
          ),
        )
      }
      if (
        [
          'manual_observation_begin',
          'manual_observation_complete',
          'manual_observation_get',
          'manual_observation_list',
          'manual_observation_delete',
          'preference_save_observation',
        ].includes(request.operation as string)
      ) {
        const scope = ['operation', 'documentId', 'projectId']
        const required =
          request.operation === 'manual_observation_list'
            ? scope
            : [
                ...scope,
                'observationId',
                ...(request.operation === 'manual_observation_begin'
                  ? ['slideId', 'shape']
                  : request.operation === 'manual_observation_complete'
                    ? ['expectedBeforeDigest', 'shape']
                    : request.operation === 'preference_save_observation'
                      ? ['text', 'expectedBeforeDigest', 'expectedAfterDigest']
                      : request.operation === 'manual_observation_delete'
                        ? ['expectedBeforeDigest', 'expectedAfterDigest']
                        : []),
              ]
        if (Object.keys(request).sort().join(',') !== required.sort().join(','))
          throw new Error('invalid_request')
        const documentId = request.documentId as string,
          projectId = request.projectId as string,
          observationId = request.observationId as string
        const bound = { documentId, projectId }
        registerLibrary(bound)
        if (request.operation === 'manual_observation_list') {
          const access = readLibrary(
            bound,
            () => manualObservations.list(documentId, projectId),
            (value) => value.length > 0,
          )
          access.assertCurrent()
          return boundedResponse({ observations: access.value })
        }
        if (typeof observationId !== 'string' || !/^[A-Za-z0-9_-]{1,80}$/.test(observationId))
          throw new Error('invalid_request')
        if (request.operation === 'manual_observation_begin') {
          const shape = parsePresentationManualObservationShape(request.shape)
          if (
            typeof request.slideId !== 'string' ||
            !request.slideId ||
            request.slideId.length > 256 ||
            Array.from(request.slideId).some(
              (c) => c.charCodeAt(0) < 32 || (c.charCodeAt(0) >= 127 && c.charCodeAt(0) <= 159),
            )
          )
            throw new Error('invalid_request')
          return boundedResponse({
            observation: manualObservations.begin(
              documentId,
              projectId,
              observationId,
              request.slideId,
              shape,
              libraryWrite(bound, false, true),
            ),
          })
        }
        const access = readLibrary(
          bound,
          () => manualObservations.get(documentId, projectId, observationId),
          (value) => !!value,
        )
        const observation = access.value
        if (request.operation === 'manual_observation_get') {
          if (!observation) throw new Error('not_found')
          access.assertCurrent()
          return boundedResponse({ observation })
        }
        if (
          request.operation === 'manual_observation_delete' ||
          request.operation === 'preference_save_observation'
        ) {
          if (
            typeof request.expectedBeforeDigest !== 'string' ||
            !/^[a-f0-9]{64}$/.test(request.expectedBeforeDigest) ||
            (request.expectedAfterDigest !== null &&
              (typeof request.expectedAfterDigest !== 'string' ||
                !/^[a-f0-9]{64}$/.test(request.expectedAfterDigest)))
          )
            throw new Error('invalid_request')
          if (
            observation &&
            (observation.before.digest !== request.expectedBeforeDigest ||
              (observation.after?.digest ?? null) !== request.expectedAfterDigest)
          )
            throw new Error('revision_conflict')
        }
        if (!observation) {
          access.assertCurrent()
          if (request.operation === 'manual_observation_delete')
            return boundedResponse({ deleted: false })
          throw new Error('not_found')
        }
        const actualScope = { documentId: observation.documentId, projectId: observation.projectId }
        if (request.operation === 'manual_observation_complete') {
          const shape = parsePresentationManualObservationShape(request.shape)
          if (shape.id !== observation.shapeId) throw new Error('invalid_request')
          if (request.expectedBeforeDigest !== observation.before.digest)
            throw new Error('revision_conflict')
          return boundedResponse({
            observation: manualObservations.complete(
              documentId,
              projectId,
              observationId,
              request.expectedBeforeDigest,
              shape,
              libraryWrite(actualScope, true),
            ),
          })
        }
        if (request.operation === 'manual_observation_delete')
          return boundedResponse({
            deleted: manualObservations.delete(
              documentId,
              projectId,
              observationId,
              libraryWrite(actualScope, true),
            ),
          })
        if (!observation.after || observation.before.digest === observation.after.digest)
          throw new Error('invalid_request')
        parseSavedPresentationPreference({
          projectId: observation.projectId,
          changeId: 'manual_' + observation.observationId,
          text: request.text,
          origin: {
            version: 1,
            observationId,
            beforeDigest: observation.before.digest,
            afterDigest: observation.after.digest,
          },
        })
        return boundedResponse({
          preference: preferenceLibrary.saveObservation(
            documentId,
            observation,
            request.text,
            libraryWrite(actualScope, true),
          ),
        })
      }
      if (request.operation === 'preference_get' || request.operation === 'preference_import') {
        const required =
          request.operation === 'preference_get'
            ? ['operation', 'documentId', 'projectId', 'changeId']
            : [
                'operation',
                'documentId',
                'projectId',
                'source',
                'expectedTextDigest',
                'approvalId',
                ...(Object.hasOwn(request, 'expectedOrigin') ? ['expectedOrigin'] : []),
              ]
        if (
          Object.keys(request).sort().join(',') !== required.sort().join(',') ||
          typeof request.documentId !== 'string' ||
          !request.documentId ||
          request.documentId.length > 2048
        )
          throw new Error('invalid_request')
        const target = { documentId: request.documentId, projectId: request.projectId as string }
        registerLibrary(target)
        if (request.operation === 'preference_get') {
          const access = readLibrary(
            target,
            () =>
              preferenceLibrary.get(
                target.documentId,
                target.projectId,
                request.changeId as string,
              ),
            (value) => !!value,
          )
          access.assertCurrent()
          return boundedResponse({ preference: access.value })
        }
        const source = parsePresentationPreferenceSource(request.source)
        if (source.documentId === target.documentId && source.projectId === target.projectId)
          throw new Error('invalid_request')
        const sourceScope = { documentId: source.documentId, projectId: source.projectId }
        const access = readLibrary(
          sourceScope,
          () => preferenceLibrary.get(source.documentId, source.projectId, source.changeId),
          (value) => !!value,
        )
        const original = access.value
        if (!original) throw new Error('not_found')
        if (
          original.reuse ||
          typeof request.approvalId !== 'string' ||
          !/^[a-f0-9]{8}-[a-f0-9]{4}-[a-f0-9]{4}-[a-f0-9]{4}-[a-f0-9]{12}$/.test(
            request.approvalId,
          ) ||
          typeof request.expectedTextDigest !== 'string' ||
          !/^[a-f0-9]{64}$/.test(request.expectedTextDigest)
        )
          throw new Error('invalid_request')
        if (createHash('sha256').update(original.text).digest('hex') !== request.expectedTextDigest)
          throw new Error('revision_conflict')
        const origin =
          request.expectedOrigin === undefined || request.expectedOrigin === null
            ? request.expectedOrigin
            : parsePresentationPreferenceOrigin(request.expectedOrigin)
        if (
          (original.origin && origin === undefined) ||
          (origin !== undefined &&
            canonicalPresentationValue(origin) !==
              canonicalPresentationValue(original.origin ?? null))
        )
          throw new Error('revision_conflict')
        access.assertCurrent()
        const targetGuard = libraryWrite(target, false, true)
        access.assertCurrent()
        const preference = preferenceLibrary.import(
          target.documentId,
          target.projectId,
          source,
          request.expectedTextDigest,
          request.approvalId,
          origin,
          {
            assertWritable: targetGuard,
            assertSourceCurrent: scopedGuard(sourceScope, access.assertCurrent),
          },
        )
        access.assertCurrent()
        targetGuard(target)
        return boundedResponse({ preference })
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
        if (request.operation === 'preference_save') {
          const input = parseSavedPresentationPreference(request.preference)
          if (input.origin !== undefined || input.reuse !== undefined)
            throw new Error('invalid_request')
          const bound = { documentId: request.documentId, projectId: input.projectId }
          registerLibrary(bound)
          return boundedResponse({
            preference: preferenceLibrary.save(
              bound.documentId,
              input,
              libraryWrite(bound, false, true),
            ),
          })
        }
        const bound = { documentId: request.documentId, projectId: request.projectId as string }
        registerLibrary(bound)
        if (request.operation === 'preference_delete') {
          const access = readLibrary(
            bound,
            () =>
              preferenceLibrary.get(bound.documentId, bound.projectId, request.changeId as string),
            (value) => !!value,
          )
          if (!access.value) {
            access.assertCurrent()
            return boundedResponse({ deleted: false })
          }
          const actualScope = { documentId: bound.documentId, projectId: access.value.projectId }
          return boundedResponse({
            deleted: preferenceLibrary.delete(
              bound.documentId,
              actualScope.projectId,
              request.changeId as string,
              libraryWrite(actualScope, true),
            ),
          })
        }
        const access = readLibrary(
          bound,
          () => preferenceLibrary.list(bound.documentId, bound.projectId),
          (value) => value.length > 0,
        )
        access.assertCurrent()
        return boundedResponse({ preferences: access.value })
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
          'package_backup_begin',
          'package_backup_chunk',
          'package_backup_finish',
          'package_backup_status',
          'package_backup_read',
          'package_backup_list',
          'package_backup_release',
          'package_backup_inventory',
        ].includes(request.operation as string)
      )
        return boundedResponse(await packageBackups(request, signal))
      if (
        [
          'master_backup_begin',
          'master_backup_chunk',
          'master_backup_finish',
          'master_backup_status',
          'master_backup_read',
          'master_backup_list',
          'master_backup_release',
          'master_backup_inventory',
        ].includes(request.operation as string)
      )
        return boundedResponse(await masterBackups(request, signal))
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
          'attachment_acquisition_history',
          'attachment_begin',
          'attachment_chunk',
          'attachment_finish',
          'attachment_extract_first_frame',
          'attachment_delete',
          'attachment_attest_license',
          'attachment_revoke_license',
          'attachment_import_url',
          'attachment_import_webpage',
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
          'accept_plan',
          'set_plan_page_lock',
          'get_plan',
          'read_import_source',
          'audit_sources',
          'read_source_audit',
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
          'production_feedback_read',
          'production_feedback_record',
          'production_feedback_compare',
        ].includes(request.operation as string)
      )
        throw new Error('invalid_request')
      const allowedKeys = [
        'production_feedback_read',
        'production_feedback_record',
        'production_feedback_compare',
      ].includes(request.operation as string)
        ? [
            'operation',
            'documentId',
            'projectId',
            'requestId',
            ...(request.operation === 'production_feedback_compare' ? ['baselineRequestId'] : []),
            ...(request.operation === 'production_feedback_record'
              ? ['expectedRevision', 'pages']
              : []),
          ]
        : request.operation === 'accept_plan'
          ? ['operation', 'documentId', 'projectId', 'decisionId', 'expectedRevision', 'planDigest']
          : ['audit_sources', 'read_source_audit'].includes(request.operation as string)
            ? ['operation', 'documentId', 'projectId', 'auditId']
            : presentationJobOperations.includes(request.operation as string)
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
                          ? ['reviewId', 'evidenceDigest', 'outcome', 'notes', 'sourceAssessment']
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
                        ? [
                            'operation',
                            'documentId',
                            'projectId',
                            'requestId',
                            'planRevision',
                            'deck',
                          ]
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
                                  : request.operation === 'read_import_source'
                                    ? [
                                        'operation',
                                        'documentId',
                                        'projectId',
                                        'requestId',
                                        'source',
                                      ]
                                    : request.operation === 'set_plan_page_lock'
                                      ? [
                                          'operation',
                                          'documentId',
                                          'projectId',
                                          'expectedRevision',
                                          'pageId',
                                          'locked',
                                        ]
                                      : request.operation === 'save_plan'
                                        ? [
                                            'operation',
                                            'documentId',
                                            'projectId',
                                            'expectedRevision',
                                            'plan',
                                          ]
                                        : request.operation === 'get_plan'
                                          ? ['operation', 'documentId', 'projectId', 'revision']
                                          : ['operation', 'documentId', 'projectId']
      const requiredKeys = allowedKeys.filter(
        (key) =>
          !(request.operation === 'compile' && ['projectId', 'planRevision'].includes(key)) &&
          !(request.operation === 'production_status' && key === 'requestId') &&
          !(request.operation === 'get_plan' && key === 'revision') &&
          !(request.operation === 'audit_sources' && key === 'auditId') &&
          !(request.operation === 'production_record_claim_review' && key === 'sourceAssessment'),
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
      if (
        request.operation === 'accept_plan' &&
        (typeof request.decisionId !== 'string' ||
          !/^[A-Za-z0-9_-]{1,128}$/.test(request.decisionId) ||
          !Number.isSafeInteger(request.expectedRevision) ||
          (request.expectedRevision as number) < 1 ||
          typeof request.planDigest !== 'string' ||
          !/^[a-f0-9]{64}$/.test(request.planDigest))
      )
        throw new Error('invalid_request')
      if (
        ['audit_sources', 'read_source_audit'].includes(request.operation as string) &&
        request.auditId !== undefined
      )
        assertPresentationId(request.auditId)
      if (request.operation === 'read_import_source') {
        assertPresentationId(request.requestId)
        if (!['compiled', 'production'].includes(request.source as string))
          throw new Error('invalid_request')
      }
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
          'production_feedback_read',
          'production_feedback_record',
          'production_feedback_compare',
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
          Number(request.offset) > MAX_PRESENTATION_SOURCE_TEXT_CHARS ||
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
          if (Object.hasOwn(request, 'sourceAssessment'))
            parsePresentationSourceAssessment(request.sourceAssessment)
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
            ...(Object.hasOwn(request, 'sourceAssessment')
              ? { sourceAssessment: request.sourceAssessment }
              : {}),
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
      if (
        request.operation === 'get_plan' &&
        request.revision !== undefined &&
        (!Number.isSafeInteger(request.revision) || Number(request.revision) < 1)
      )
        throw new Error('invalid_request')
      if (request.operation === 'save_plan') {
        if (!Number.isSafeInteger(request.expectedRevision) || Number(request.expectedRevision) < 0)
          throw new Error('invalid_request')
        try {
          plan = parsePresentationPlan(request.plan)
        } catch {
          throw new Error('invalid_plan')
        }
      }
      if (request.operation === 'production_feedback_compare') {
        assertPresentationId(request.baselineRequestId)
        if (request.baselineRequestId === request.requestId) throw new Error('invalid_request')
      }
      if (request.operation === 'production_feedback_record') {
        if (!Number.isSafeInteger(request.expectedRevision) || Number(request.expectedRevision) < 0)
          throw new Error('invalid_request')
        parsePresentationProductionFeedbackPages(request.pages)
      }
      const projectId = request.projectId ?? deck?.id
      assertPresentationId(projectId)
      if (plan && plan.projectId !== projectId) throw new Error('invalid_plan')
      if (deck && deck.id !== projectId) throw new Error('invalid_request')
      const key = `${resolve(options.userDataPath)}\0${projectId}`
      const ordinaryProject =
        ordinaryProjectReadOperations.has(request.operation as string) ||
        ordinaryProjectWriteOperations.has(request.operation as string)
      foregroundWork = registerPresentationProjectWork({
        scope: { root: options.userDataPath, projectId, documentId },
        signal,
      })
      signal = foregroundWork.signal
      const readonlyProject =
        readonlyProductionOperations.has(request.operation as string) ||
        ordinaryProjectReadOperations.has(request.operation as string)
      const existingProject = ordinaryProject
        ? store.projectScope(projectId, documentId)
        : undefined
      const control = ordinaryProject ? lifecycleStore.read({ projectId, documentId }) : undefined
      if (control) lifecycleStore.assertActive({ projectId, documentId }, control.revision)
      if (ordinaryProject && readonlyProject && !existingProject && !control)
        throw new Error(
          request.operation === 'get_plan' && request.revision !== undefined
            ? 'plan_revision_unavailable'
            : 'not_found',
        )
      const creating =
        ordinaryProject &&
        !existingProject &&
        ['compile', 'save_plan'].includes(request.operation as string)
      if (
        creating &&
        (request.operation === 'save_plan'
          ? request.expectedRevision !== 0
          : request.planRevision !== undefined)
      )
        throw new Error('revision_conflict')
      if (ordinaryProject && !existingProject && !control && !creating) throw new Error('not_found')
      const projectLease = creating
        ? capturePresentationProjectCreationLease({
            store: lifecycleStore,
            scope: { projectId, documentId },
            signal,
          })
        : (request.operation as string).startsWith('production_') || ordinaryProject
          ? (readonlyProject
              ? capturePresentationProjectReadLease
              : capturePresentationProjectWriteLease)({
              store: lifecycleStore,
              scope: { projectId, documentId },
              readExistingProject: (scope) => store.projectScope(scope.projectId, scope.documentId),
              ...(presentationJobOperations.includes(request.operation as string)
                ? {}
                : { signal }),
            })
          : undefined
      const assertProjectCurrent = projectLease
        ? 'assertCurrent' in projectLease
          ? projectLease.assertCurrent
          : projectLease.assertWritable
        : undefined
      const assertProjectWrite = readonlyProject
        ? () => {
            throw new Error('access_denied')
          }
        : assertProjectCurrent
      const release = await acquireProjectLock(options.userDataPath, projectId, signal)
      const guardedResponse = (value: unknown) => {
        assertProjectCurrent?.()
        return boundedResponse(value)
      }
      try {
        checkAbort(signal)
        assertProjectCurrent?.()
        if (request.operation === 'production_feedback_compare') {
          const baseline = store.production(
            projectId,
            documentId,
            request.baselineRequestId as string,
          )
          const candidate = store.production(projectId, documentId, request.requestId as string)
          if (!baseline || !candidate) throw new Error('not_found')
          if (
            [baseline, candidate].some((task) =>
              task.pages.some((page) => page.state !== 'compiled'),
            )
          )
            throw new Error('page_not_ready')
          try {
            const side = (task: NonNullable<typeof baseline>) => ({
              requestId: task.requestId,
              inputDigest: task.inputDigest,
              planDigest: task.planDigest,
              planRevision: task.plan.revision,
              plan: parsePresentationPlan(task.plan.plan),
              feedback: store.productionFeedback(projectId, documentId, task.requestId) ?? null,
            })
            return guardedResponse({
              comparison: buildPresentationFeedbackComparison({
                projectId,
                documentId,
                baseline: side(baseline),
                candidate: side(candidate),
              }),
            })
          } catch (error) {
            if (
              error instanceof Error &&
              error.message === 'presentation_feedback_comparison_invalid'
            )
              throw new Error('invalid_state', { cause: error })
            throw error
          }
        }
        if (request.operation === 'production_feedback_read')
          return guardedResponse({
            feedback:
              store.productionFeedback(projectId, documentId, request.requestId as string) ?? null,
          })
        if (request.operation === 'production_feedback_record') {
          assertProjectCurrent?.()
          return guardedResponse({
            feedback: store.recordProductionFeedback(
              projectId,
              documentId,
              request.requestId as string,
              request.expectedRevision as number,
              request.pages,
            ),
          })
        }
        if (presentationJobOperations.includes(request.operation as string)) {
          if (
            ['production_job_start', 'production_job_resume'].includes(request.operation as string)
          ) {
            const frozen = store.production(projectId, documentId, request.requestId as string)
            if (frozen)
              await readBoundPresentationResearch(
                parsePresentationPlan(frozen.plan.plan),
                documentId,
                projectId,
                (ledgerId) => researchStore.read(documentId, projectId, ledgerId),
                signal,
              )
          }
          checkAbort(signal)
          assertProjectCurrent?.()
          const response = encode(
            handlePresentationJob(key, request, {
              store,
              compile,
              attachments,
              assertWritable: assertProjectWrite,
              readResearch: (ledgerId) => researchStore.read(documentId, projectId, ledgerId),
            }),
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
        ) {
          const result = await handlePresentationDeliveryReport(
            request,
            store,
            attachments,
            signal,
            (ledgerId) => researchStore.read(documentId, projectId, ledgerId),
            assertProjectWrite,
          )
          assertProjectCurrent?.()
          return guardedResponse(result)
        }
        if ((request.operation as string).startsWith('production_')) {
          const result = await handlePresentationProduction(
            request,
            {
              store,
              compile,
              attachments,
              assertWritable: assertProjectWrite,
              readResearch: (ledgerId) => researchStore.read(documentId, projectId, ledgerId),
            },
            signal,
          )
          assertProjectCurrent?.()
          return guardedResponse(result)
        }
        if (request.operation === 'read_import_source')
          return guardedResponse(
            readPresentationImportSource(
              store,
              projectId,
              documentId,
              request.requestId as string,
              request.source as 'compiled' | 'production',
            ),
          )
        if (request.operation === 'accept_plan') {
          assertProjectWrite?.()
          const acceptance = store.acceptPlan(
            projectId,
            documentId,
            request.decisionId as string,
            request.expectedRevision as number,
            request.planDigest as string,
          )
          return guardedResponse({ projectId, documentId, acceptance })
        }
        if (request.operation === 'set_plan_page_lock') {
          assertProjectWrite?.()
          const record = store.setPlanPageLock(
            projectId,
            documentId,
            request.expectedRevision as number,
            request.pageId as string,
            request.locked as boolean,
          )
          return guardedResponse({
            projectId,
            revision: record.revision,
            plan: parsePresentationPlan(record.plan),
          })
        }
        if (request.operation === 'save_plan' || request.operation === 'get_plan') {
          if (request.operation === 'save_plan') {
            await readBoundPresentationResearch(
              plan!,
              documentId,
              projectId,
              (ledgerId) => researchStore.read(documentId, projectId, ledgerId),
              signal,
            )
            checkAbort(signal)
            assertProjectCurrent?.()
            const previousPlan = store.plan(projectId, documentId)
            if (previousPlan && previousPlan.revision === request.expectedRevision) {
              try {
                assertBrandKitRevision(parsePresentationPlan(previousPlan.plan), plan!)
              } catch {
                throw new Error('invalid_plan')
              }
            }
          }
          if (request.operation === 'save_plan') assertProjectWrite?.()
          const record =
            request.operation === 'save_plan'
              ? store.savePlan(projectId, documentId, request.expectedRevision as number, plan)
              : request.revision === undefined
                ? store.plan(projectId, documentId)
                : store.planRevision(projectId, documentId, request.revision as number)
          if (!record)
            throw new Error(
              request.revision === undefined ? 'not_found' : 'plan_revision_unavailable',
            )
          return guardedResponse({
            projectId,
            revision: record.revision,
            plan: parsePresentationPlan(record.plan),
          })
        }
        if (['audit_sources', 'read_source_audit'].includes(request.operation as string))
          return guardedResponse(
            await handlePresentationSourceAudit(
              request,
              store,
              attachments,
              signal,
              assertProjectWrite,
            ),
          )
        if (request.operation === 'status') {
          let researchStatus: Record<string, unknown>
          try {
            researchStatus = { researchSummary: await researchStore.summary(documentId, projectId) }
          } catch {
            researchStatus = { researchHistoryUnavailable: true }
          }
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
          let sourceAuditStatus: Record<string, unknown>
          let acceptanceStatus: Record<string, unknown>
          try {
            acceptanceStatus = { planAcceptance: store.planAcceptances(projectId, documentId) }
          } catch {
            acceptanceStatus = { planAcceptanceUnavailable: true }
          }
          try {
            sourceAuditStatus = {
              sourceAuditHistory: presentationSourceAuditHistory(
                store.sourceAudits(projectId, documentId),
              ),
            }
          } catch {
            sourceAuditStatus = { sourceAuditHistoryUnavailable: true }
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
          let assetStatus: Record<string, unknown> = {}
          if (productionRecord) {
            try {
              assetStatus = {
                assetHistory: store.productionAssets(
                  projectId,
                  documentId,
                  productionRecord.requestId,
                ),
              }
            } catch {
              assetStatus = { assetHistoryUnavailable: true }
            }
          }
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
                  | 'source_mismatch'
                  | 'locator_mismatch'
              }[]
            | undefined
          let sourcePreparationUnavailable = false
          if (plan) {
            const references = plan.value.sources.flatMap((source) => {
              const attachmentId = presentationSourceAttachmentId(source)
              return attachmentId
                ? [
                    {
                      sourceId: source.id,
                      attachmentId,
                      excerpt: source.excerpt,
                      uri: source.uri,
                      locator: source.locator,
                    },
                  ]
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
                      sourceUrlHash?: string
                    }
                    if (item.attachmentId !== reference.attachmentId)
                      throw new Error('invalid_state')
                    let status: NonNullable<typeof sourcePreparation>[number]['status'] =
                      item.status === 'ready' ? 'unsupported' : item.status
                    if (!matchesFetchedSourceUrl(reference.uri, item.sourceUrlHash)) {
                      status = 'source_mismatch'
                    } else if (item.status === 'ready' && item.kind === 'text') {
                      const match = (await attachments(
                        {
                          operation: 'attachment_match_excerpt',
                          documentId,
                          attachmentId: reference.attachmentId,
                          excerpt: reference.excerpt,
                          ...(canonicalSourceLocator(reference.locator)
                            ? { locator: canonicalSourceLocator(reference.locator) }
                            : {}),
                        },
                        signal,
                      )) as { attachmentId?: unknown; status?: unknown; locator?: unknown }
                      if (match.attachmentId !== reference.attachmentId)
                        throw new Error('invalid_state')
                      const matchStatus = {
                        found: 'excerpt_matched',
                        not_found: 'excerpt_mismatch',
                        empty_excerpt: 'excerpt_missing',
                        not_ready: 'uploading',
                        unsupported: 'unsupported',
                      }[String(match.status)] as
                        NonNullable<typeof sourcePreparation>[number]['status'] | undefined
                      if (!matchStatus) throw new Error('invalid_state')
                      status = matchStatus
                      if (
                        matchStatus === 'excerpt_matched' &&
                        canonicalSourceLocator(reference.locator) &&
                        typeof match.locator === 'string' &&
                        canonicalSourceLocator(reference.locator) !== match.locator
                      )
                        status = 'locator_mismatch'
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
            return guardedResponse({
              projectId,
              ...(control ? { createdAt: control.createdAt } : {}),
              title: plan.value.title,
              status: 'planned',
              deliveryBundlesAvailable: true,
              researchAvailable: true,
              ...researchStatus,
              ...(production ? { production, productionTasks } : {}),
              slideCount: plan.value.slides.length,
              slides: plan.value.slides.map(({ id, title }) => ({ id, title })),
              history: [],
              plan,
              ...commentStatus,
              ...sourceAuditStatus,
              ...acceptanceStatus,
              ...assetStatus,
              ...preparationStatus,
            })
          }
          const latestDeck = savedDeck(latest.deck)
          const compiled = store.latest(projectId, documentId)
          const checks =
            latest.status === 'compiled'
              ? (latest.result as { report: PresentationCompileReport }).report.checks
              : undefined
          return guardedResponse({
            projectId,
            ...(control ? { createdAt: control.createdAt } : {}),
            title: latestDeck.title,
            status: latest.status,
            deliveryBundlesAvailable: true,
            researchAvailable: true,
            ...researchStatus,
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
            ...sourceAuditStatus,
            ...acceptanceStatus,
            ...assetStatus,
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
          return guardedResponse(record.result)
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
          if (binding)
            await readBoundPresentationResearch(
              parsePresentationPlan(binding.plan),
              documentId,
              projectId,
              (ledgerId) => researchStore.read(documentId, projectId, ledgerId),
              signal,
            )
          checkAbort(signal)
          assertProjectWrite?.()
          record = store.begin(projectId, documentId, requestId, deck, binding)
        }
        if (!record) throw new Error('not_found')
        if (record.status === 'compiled') return guardedResponse(record.result)
        if (request.operation === 'resume' && record.plan)
          await readBoundPresentationResearch(
            parsePresentationPlan(record.plan.plan),
            documentId,
            projectId,
            (ledgerId) => researchStore.read(documentId, projectId, ledgerId),
            signal,
          )
        checkAbort(signal)
        assertProjectCurrent?.()
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
          for (const slide of inputDeck.slides) {
            await assertCitedPresentationSourcesReady(
              plan,
              slide,
              documentId,
              attachments,
              signal,
              sourceReadiness,
            )
            assertProjectCurrent?.()
          }
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
              assertProjectCurrent?.()
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
        assertProjectCurrent?.()
        if (record.plan) assertBrandLogoAsset(parsePresentationPlan(record.plan.plan), assets)
        // Only attachment_asset can add evidence after document-bound digest validation.
        const compiled = await compile({ ...inputDeck, assets }, { trustedAssetEvidence: true })
        checkAbort(signal)
        assertProjectCurrent?.()
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
        assertProjectWrite?.()
        store.complete(record, result)
        return response
      } catch (error) {
        assertProjectCurrent?.()
        throw error
      } finally {
        release()
      }
    } catch (error) {
      let code = 'compile_failed'
      try {
        const message = error instanceof Error ? error.message : undefined
        if (typeof message === 'string' && errorCodes.has(message)) code = message
      } catch {
        /* Only finite error codes are sent to the paired client. */
      }
      return encode({ error: code })
    } finally {
      foregroundWork?.finish()
    }
  }
}

export { acquireProjectLock as acquirePresentationProjectLock }
