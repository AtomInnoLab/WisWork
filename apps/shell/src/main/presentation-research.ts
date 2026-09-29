import { PresentationResearchStore } from '@wiswork/project-store/presentation-research-store'
import {
  parsePresentationResearchDraft,
  type PresentationResearchEvidence,
  type PresentationResearchRecord,
} from '@wiswork/project-store/presentation-research'
import { createPresentationAttachmentService } from './presentation-attachments'
import { canonicalSourceLocator, matchesFetchedSourceUrl } from './presentation-source-audit'
export function createPresentationResearchService(options: {
  userDataPath: string
  attachments?: ReturnType<typeof createPresentationAttachmentService>
}) {
  const store = new PresentationResearchStore(options.userDataPath),
    attachments = options.attachments ?? createPresentationAttachmentService(options)
  async function evidence(
    documentId: string,
    source: ReturnType<typeof parsePresentationResearchDraft>['sources'][number],
    signal: AbortSignal,
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
      details = (await attachments(
        { operation: 'attachment_metadata', documentId, attachmentId },
        signal,
      )) as Record<string, unknown>
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
    const result = (await attachments(
      {
        operation: 'attachment_match_excerpt',
        documentId,
        attachmentId,
        excerpt: source.excerpt,
        ...(locator ? { locator } : {}),
      },
      signal,
    )) as {
      attachmentId: string
      status: PresentationResearchEvidence['status']
      offset?: number
      locator?: string
    }
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
        research_build: ['projectId', 'ledgerId', 'expectedRevision', 'draft'],
      }
    if (typeof op !== 'string' || !Object.hasOwn(fields, op)) throw new Error('invalid_request')
    const allowed = ['operation', 'documentId', ...fields[op]!]
    if (
      Object.keys(request).some((k) => !allowed.includes(k)) ||
      allowed.some((k) => !Object.hasOwn(request, k)) ||
      typeof request.documentId !== 'string' ||
      !request.documentId.trim() ||
      request.documentId.length > 4096 ||
      Buffer.byteLength(JSON.stringify(request)) > 300 * 1024
    )
      throw new Error('invalid_request')
    if (signal.aborted) throw new Error('aborted')
    if (op === 'research_capabilities') return { version: 1, available: true }
    const projectId = request.projectId
    if (typeof projectId !== 'string' || !/^[A-Za-z0-9_-]{1,128}$/.test(projectId))
      throw new Error('invalid_request')
    const documentId = request.documentId
    if (op === 'research_latest') {
      const record = await store.latestCompleted(documentId, projectId)
      if (signal.aborted) throw new Error('aborted')
      return { record }
    }
    if (op === 'research_list') {
      const h = await store.summary(documentId, projectId)
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
    if (
      typeof request.expectedRevision !== 'number' ||
      !Number.isSafeInteger(request.expectedRevision) ||
      request.expectedRevision < 0
    )
      throw new Error('invalid_request')
    const draft = parsePresentationResearchDraft(request.draft)
    const begin = await store.begin(documentId, projectId, request.expectedRevision, id, draft)
    if (!begin.created)
      return { history: await store.summary(documentId, projectId), record: begin.record }
    let record: PresentationResearchRecord
    const sources: PresentationResearchEvidence[] = []
    try {
      for (const source of draft.sources) sources.push(await evidence(documentId, source, signal))
      if (signal.aborted) throw new Error('aborted')
      record = await store.finish(documentId, projectId, id, { state: 'completed', sources })
    } catch (e) {
      const code = e instanceof Error ? e.message : ''
      record = await store.finish(documentId, projectId, id, {
        state: 'failed',
        error:
          code === 'aborted'
            ? 'aborted'
            : code === 'source_unavailable'
              ? 'source_unavailable'
              : 'invalid_state',
        sources,
      })
    }
    return { history: await store.summary(documentId, projectId), record }
  }
}
