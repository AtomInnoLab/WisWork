import { registerPresentationProjectWork } from './presentation-project-work'
import type {
  PresentationProjectReadLease,
  PresentationProjectWriteLease,
} from './presentation-project-write-lease'
import { PresentationResearchStore } from '@wiswork/project-store/presentation-research-store'
import {
  parsePresentationResearchDraft,
  type PresentationResearchEvidence,
  type PresentationResearchRecord,
} from '@wiswork/project-store/presentation-research'
import { createPresentationAttachmentService } from './presentation-attachments'
import { canonicalSourceLocator, matchesFetchedSourceUrl } from './presentation-source-audit'
export function createPresentationResearchService(optionsValue: {
  userDataPath: string
  attachments?: ReturnType<typeof createPresentationAttachmentService>
  acquireProjectLock?: (projectId: string, signal?: AbortSignal) => Promise<() => void>
  captureProjectLease?: (input: {
    scope: Readonly<{ documentId: string; projectId: string }>
    operation: string
    signal: AbortSignal
  }) =>
    | PresentationProjectReadLease
    | PresentationProjectWriteLease
    | Promise<PresentationProjectReadLease | PresentationProjectWriteLease>
  assertRecordUnprotected?: (
    documentId: string,
    projectId: string,
    ledgerId: string,
    signal: AbortSignal,
  ) => Promise<void>
}) {
  const options = { ...optionsValue }
  const store = new PresentationResearchStore(options.userDataPath),
    attachments = options.attachments ?? createPresentationAttachmentService(options)
  async function evidence(
    documentId: string,
    source: ReturnType<typeof parsePresentationResearchDraft>['sources'][number],
    signal: AbortSignal,
    assertCurrent: () => void,
  ): Promise<PresentationResearchEvidence> {
    if (signal.aborted) throw new Error('aborted')
    const attachmentId =
      /^attachment:([a-f0-9]{64})$/.exec(source.uri)?.[1] ?? source.snapshotAttachmentId
    if (!attachmentId) return { sourceId: source.id, status: 'missing', provenance: 'unavailable' }
    const unavailable = (
      status: PresentationResearchEvidence['status'],
    ): PresentationResearchEvidence => ({
      sourceId: source.id,
      attachmentId,
      status,
      provenance: 'unavailable',
    })
    let details: Record<string, unknown>
    try {
      details = structuredClone(
        await attachments({ operation: 'attachment_metadata', documentId, attachmentId }, signal),
      ) as Record<string, unknown>
      assertCurrent()
    } catch (e) {
      if (e instanceof Error && e.message === 'not_found') return unavailable('missing')
      throw e
    }
    if (
      details.attachmentId !== attachmentId ||
      details.sha256 !== attachmentId ||
      !['ready', 'uploading', 'failed'].includes(details.status as string)
    )
      throw new Error('invalid_state')
    if (details.status !== 'ready') return unavailable('not_ready')
    if (details.kind !== 'text') return unavailable('unsupported')
    const fetched = details.sourceUrlHash !== undefined
    if (fetched && !matchesFetchedSourceUrl(source.uri, details.sourceUrlHash))
      return unavailable('source_mismatch')
    const locator = canonicalSourceLocator(source.locator)
    const result = structuredClone(
      await attachments(
        {
          operation: 'attachment_match_excerpt',
          documentId,
          attachmentId,
          excerpt: source.excerpt,
          ...(locator ? { locator } : {}),
        },
        signal,
      ),
    ) as {
      attachmentId: string
      status: PresentationResearchEvidence['status']
      offset?: number
      locator?: string
    }
    assertCurrent()
    if (
      result.attachmentId !== attachmentId ||
      !['found', 'not_found', 'empty_excerpt', 'not_ready', 'unsupported'].includes(
        result.status,
      ) ||
      (result.status === 'found'
        ? !Number.isSafeInteger(result.offset) ||
          result.offset! < 0 ||
          result.offset! > 8_000_000 ||
          (result.locator !== undefined && !/^第 [1-9]\d{0,5} (页|段)$/.test(result.locator))
        : result.offset !== undefined || result.locator !== undefined)
    )
      throw new Error('invalid_state')
    if (['not_ready', 'unsupported'].includes(result.status)) return unavailable(result.status)
    const status =
      result.status === 'found' && locator && result.locator !== locator
        ? 'not_found'
        : result.status
    const common = {
      sourceId: source.id,
      attachmentId,
      status,
      sha256: attachmentId,
      ...(status === 'found'
        ? { offset: result.offset, ...(result.locator ? { locator: result.locator } : {}) }
        : {}),
    }
    if (fetched) {
      if (
        typeof details.retrievedAt !== 'number' ||
        !Number.isSafeInteger(details.retrievedAt) ||
        details.retrievedAt < 1 ||
        !source.snapshotAttachmentId
      )
        throw new Error('invalid_state')
      return {
        ...common,
        provenance: 'fetched_url_matched',
        retrievedAt: new Date(details.retrievedAt).toISOString(),
      }
    }
    return { ...common, provenance: 'user_supplied' }
  }
  return async (request: Record<string, unknown>, signal: AbortSignal): Promise<unknown> => {
    const op = request.operation,
      fields: Record<string, string[]> = {
        research_capabilities: [],
        research_list: ['projectId'],
        research_latest: ['projectId'],
        research_read: ['projectId', 'ledgerId'],
        research_abandon: ['projectId', 'ledgerId', 'expectedDraftDigest', 'expectedRevision'],
        research_build: ['projectId', 'ledgerId', 'expectedRevision', 'draft'],
        research_delete: [
          'projectId',
          'ledgerId',
          'deleteId',
          'expectedDraftDigest',
          'expectedRevision',
        ],
        research_delete_status: ['projectId', 'deleteId'],
      }
    if (typeof op !== 'string' || !Object.hasOwn(fields, op)) throw new Error('invalid_request')
    const required = ['operation', 'documentId', ...fields[op]!]
    const optional =
      op === 'research_capabilities'
        ? ['includeCleanup', 'includeRecovery']
        : ['research_list', 'research_build'].includes(op)
          ? ['historyVersion']
          : []
    const allowed = [...required, ...optional]
    if (
      Object.keys(request).some((k) => !allowed.includes(k)) ||
      required.some((k) => !Object.hasOwn(request, k)) ||
      (Object.hasOwn(request, 'includeCleanup') && request.includeCleanup !== true) ||
      (Object.hasOwn(request, 'includeRecovery') &&
        (request.includeRecovery !== true || request.includeCleanup !== true)) ||
      (Object.hasOwn(request, 'historyVersion') && request.historyVersion !== 2) ||
      typeof request.documentId !== 'string' ||
      !request.documentId.trim() ||
      request.documentId.length > 4096 ||
      Buffer.byteLength(JSON.stringify(request)) > 300 * 1024
    )
      throw new Error('invalid_request')
    if (signal.aborted) throw new Error('aborted')
    if (op === 'research_capabilities')
      return request.includeRecovery
        ? {
            version: 1,
            available: true,
            cleanupAvailable: true,
            recoveryAvailable: true,
            historyVersions: [1, 2],
          }
        : request.includeCleanup
          ? { version: 1, available: true, cleanupAvailable: true, historyVersions: [1, 2] }
          : { version: 1, available: true }
    request = structuredClone(request)
    const projectId = request.projectId
    if (typeof projectId !== 'string' || !/^[A-Za-z0-9_-]{1,128}$/.test(projectId))
      throw new Error('invalid_request')
    const documentId = request.documentId
    if (typeof documentId !== 'string' || !documentId.trim() || documentId.length > 4096)
      throw Error('invalid_request')
    const validIdentifier = (value: unknown) =>
      typeof value === 'string' && /^[A-Za-z0-9_-]{1,128}$/.test(value)
    if (
      (['research_read', 'research_build', 'research_delete', 'research_abandon'].includes(op) &&
        !validIdentifier(request.ledgerId)) ||
      (['research_delete', 'research_delete_status'].includes(op) &&
        !validIdentifier(request.deleteId)) ||
      (['research_build', 'research_delete', 'research_abandon'].includes(op) &&
        (!Number.isSafeInteger(request.expectedRevision) ||
          Number(request.expectedRevision) < 0)) ||
      (['research_delete', 'research_abandon'].includes(op) &&
        (typeof request.expectedDraftDigest !== 'string' ||
          !/^[a-f0-9]{64}$/.test(request.expectedDraftDigest)))
    )
      throw Error('invalid_request')
    const capturedDraft =
      op === 'research_build' ? parsePresentationResearchDraft(request.draft) : undefined
    const scope = Object.freeze({ documentId, projectId })
    const work = options.captureProjectLease
      ? registerPresentationProjectWork({ scope: { root: options.userDataPath, ...scope }, signal })
      : undefined
    signal = work?.signal ?? signal
    let lease: PresentationProjectReadLease | PresentationProjectWriteLease | undefined
    const assertCurrent = () => {
      if (!options.captureProjectLease) return
      if (signal.aborted) throw Error('aborted')
      if (lease) {
        if ('assertWritable' in lease) lease.assertWritable()
        else lease.assertCurrent()
      }
      if (signal.aborted) throw Error('aborted')
    }
    const assertWritable = () => {
      assertCurrent()
      if (lease && !('assertWritable' in lease)) throw Error('access_denied')
    }
    let projectLock: Promise<(() => void) | undefined> | undefined
    try {
      const captured = options.captureProjectLease?.({ scope, operation: op, signal })
      lease = captured instanceof Promise ? await captured : captured
      assertCurrent()
      projectLock = options.acquireProjectLock?.(projectId, signal)
      const run = async () => {
        const release = await projectLock
        try {
          assertCurrent()
          if (signal.aborted) throw new Error('aborted')
          const summary = async () => {
            const result = await store.summary(documentId, projectId)
            assertCurrent()
            if (result.version === 2 && request.historyVersion !== 2)
              throw new Error('upgrade_required')
            return result
          }
          if (op === 'research_delete_status') {
            if (
              typeof request.deleteId !== 'string' ||
              !/^[A-Za-z0-9_-]{1,128}$/.test(request.deleteId)
            )
              throw new Error('invalid_request')
            const receipt = await store.deletedReceipt(documentId, projectId, request.deleteId)
            if (signal.aborted) throw new Error('aborted')
            if (!receipt) throw new Error('not_found')
            return receipt
          }
          if (op === 'research_latest') {
            const record = await store.latestCompleted(documentId, projectId)
            if (signal.aborted) throw new Error('aborted')
            return { record }
          }
          if (op === 'research_list') {
            const h = await summary()
            if (signal.aborted) throw new Error('aborted')
            return h
          }
          const id = request.ledgerId
          if (typeof id !== 'string' || !/^[A-Za-z0-9_-]{1,128}$/.test(id))
            throw new Error('invalid_request')
          if (op === 'research_read') {
            const r = await store.read(documentId, projectId, id)
            if (signal.aborted) throw new Error('aborted')
            return r
          }
          if (op === 'research_delete') {
            if (
              typeof request.deleteId !== 'string' ||
              !/^[A-Za-z0-9_-]{1,128}$/.test(request.deleteId) ||
              typeof request.expectedDraftDigest !== 'string' ||
              !/^[a-f0-9]{64}$/.test(request.expectedDraftDigest) ||
              !Number.isSafeInteger(request.expectedRevision) ||
              Number(request.expectedRevision) < 0
            )
              throw new Error('invalid_request')
            const receipt = await store
              .deletedReceipt(documentId, projectId, request.deleteId)
              .catch((error) => {
                if (error instanceof Error && error.message === 'not_found') return undefined
                throw error
              })
            if (receipt)
              return await store.deleteRecord(
                documentId,
                projectId,
                request.expectedRevision as number,
                request.deleteId,
                id,
                request.expectedDraftDigest,
                signal,
                options.captureProjectLease ? assertWritable : undefined,
              )
            const record = await store.read(documentId, projectId, id)
            if (!record) throw new Error('not_found')
            if (record.state === 'running') throw new Error('record_running')
            if (!options.assertRecordUnprotected || !options.acquireProjectLock)
              throw new Error('invalid_state')
            await options.assertRecordUnprotected(documentId, projectId, id, signal)
            assertCurrent()
            if (signal.aborted) throw new Error('aborted')
            return await store.deleteRecord(
              documentId,
              projectId,
              request.expectedRevision as number,
              request.deleteId,
              id,
              request.expectedDraftDigest,
              signal,
              options.captureProjectLease ? assertWritable : undefined,
            )
          }
          if (op === 'research_abandon') {
            if (!options.acquireProjectLock) throw new Error('invalid_state')
            if (
              !Number.isSafeInteger(request.expectedRevision) ||
              Number(request.expectedRevision) < 0 ||
              typeof request.expectedDraftDigest !== 'string' ||
              !/^[a-f0-9]{64}$/.test(request.expectedDraftDigest)
            )
              throw new Error('invalid_request')
            return await store.abandon(
              documentId,
              projectId,
              request.expectedRevision as number,
              id,
              request.expectedDraftDigest,
              signal,
              options.captureProjectLease ? assertWritable : undefined,
            )
          }
          if (
            typeof request.expectedRevision !== 'number' ||
            !Number.isSafeInteger(request.expectedRevision) ||
            request.expectedRevision < 0
          )
            throw new Error('invalid_request')
          const draft = capturedDraft!
          await summary()
          if (signal.aborted) throw new Error('aborted')
          const begin = await store.begin(
            documentId,
            projectId,
            request.expectedRevision,
            id,
            draft,
            options.captureProjectLease ? assertWritable : undefined,
          )
          if (!begin.created) return { history: await summary(), record: begin.record }
          let record: PresentationResearchRecord
          const sources: PresentationResearchEvidence[] = []
          try {
            for (const source of draft.sources)
              sources.push(await evidence(documentId, source, signal, assertCurrent))
            if (signal.aborted) throw new Error('aborted')
            assertWritable()
            record = await store.finish(
              documentId,
              projectId,
              id,
              { state: 'completed', sources },
              options.captureProjectLease ? assertWritable : undefined,
            )
          } catch (e) {
            const code = e instanceof Error ? e.message : ''
            assertWritable()
            record = await store.finish(
              documentId,
              projectId,
              id,
              {
                state: 'failed',
                error:
                  code === 'aborted'
                    ? 'aborted'
                    : code === 'source_unavailable'
                      ? 'source_unavailable'
                      : 'invalid_state',
                sources,
              },
              options.captureProjectLease ? assertWritable : undefined,
            )
          }
          return { history: await summary(), record }
        } finally {
          release?.()
          projectLock = undefined
        }
      }
      const result = await run()
      assertCurrent()
      return result
    } finally {
      // Admission may fail while queued; settle and release the acquired lock before finishing Work.
      try {
        if (projectLock) (await projectLock)?.()
      } finally {
        work?.finish()
      }
    }
  }
}
