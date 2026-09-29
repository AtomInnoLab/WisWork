import { createHash } from 'node:crypto'
import { resolve } from 'node:path'
import {
  PresentationStore,
  PresentationLifecycleStore,
  PRESENTATION_LIFECYCLE_RESOURCE_KINDS,
  type PresentationLifecycleScope,
  type PresentationLifecycleRecord,
  type PresentationLifecycleResourceKind,
} from '@wiswork/project-store'
import {
  inspectPresentationProjectInventory,
  type PresentationInventoryResource,
} from './presentation-project-inventory'
import { removePresentationProjectDeletionResource } from './presentation-project-deletion-resources'
import { stopPresentationProjectWork } from './presentation-project-work'
const check = (signal?: AbortSignal) => {
  if (signal?.aborted) throw Error('aborted')
}
function owned<T>(value: T, keys: string[]): T {
  if (
    !value ||
    typeof value !== 'object' ||
    Array.isArray(value) ||
    Object.keys(value).sort().join(',') !== [...keys].sort().join(',')
  )
    throw Error('invalid_request')
  if (Buffer.byteLength(JSON.stringify(value)) > 32 * 1024) throw Error('invalid_request')
  return structuredClone(value)
}
function scope(value: PresentationLifecycleScope) {
  const s = owned(value, ['projectId', 'documentId'])
  if (
    !s ||
    Object.keys(s).sort().join(',') !== 'documentId,projectId' ||
    typeof s.projectId !== 'string' ||
    !/^[A-Za-z0-9_-]{1,128}$/.test(s.projectId) ||
    typeof s.documentId !== 'string' ||
    !s.documentId.trim() ||
    s.documentId.length > 4096 ||
    Array.from(s.documentId).some(
      (c) => c.charCodeAt(0) < 32 || (c.charCodeAt(0) >= 127 && c.charCodeAt(0) <= 159),
    )
  )
    throw Error('invalid_request')
  return Object.freeze(s)
}
function revision(value: unknown, nullable = false) {
  if (nullable && value === null) return
  if (typeof value !== 'number' || !Number.isSafeInteger(value) || value < 0)
    throw Error('invalid_request')
}
const known = (
  r: PresentationInventoryResource,
): r is PresentationInventoryResource & { kind: PresentationLifecycleResourceKind } =>
  PRESENTATION_LIFECYCLE_RESOURCE_KINDS.includes(r.kind as PresentationLifecycleResourceKind)
const content = (resources: PresentationInventoryResource[]) =>
  resources.filter((r) => r.kind !== 'lifecycle_control')
const token = (
  s: PresentationLifecycleScope,
  rev: number | null,
  resources: PresentationInventoryResource[],
) =>
  createHash('sha256')
    .update(JSON.stringify({ scope: s, revision: rev, resources: content(resources) }))
    .digest('hex')
/** Internal only; callers must inject the actual lock shared by all project writers. */
export function createPresentationProjectDeletionService(optionsValue: {
  userDataPath: string
  acquireProjectLock: (projectId: string) => Promise<() => void>
}) {
  const options = { ...optionsValue, userDataPath: resolve(optionsValue.userDataPath) },
    store = new PresentationStore(options.userDataPath),
    life = new PresentationLifecycleStore(options.userDataPath)
  const inventory = (s: PresentationLifecycleScope, signal?: AbortSignal) =>
    inspectPresentationProjectInventory({ userDataPath: options.userDataPath, ...s, signal })
  const current = (s: PresentationLifecycleScope, rev: number, id: string) => {
    const r = life.read(s)
    if (!r || r.revision !== rev) throw Error('revision_conflict')
    if (r.state !== 'deleting' || r.deletion?.deletionId !== id) throw Error('deletion_conflict')
    return r
  }
  const report = (r: PresentationLifecycleRecord, resources: PresentationInventoryResource[]) => ({
    state: r.state === 'deleted' ? ('deleted' as const) : ('partial' as const),
    revision: r.revision,
    deletionId: r.deletion!.deletionId,
    projectContentRetained: content(resources).some((v) => v.kind === 'project'),
    counts: {
      removed: r.deletion!.resources.filter(
        (v) => v.status === 'removed' || v.status === 'reference_removed',
      ).length,
      pending: r.deletion!.resources.filter((v) => v.status === 'pending').length,
      failed: r.deletion!.resources.filter((v) => v.status === 'failed').length,
      retained: content(resources).length,
    },
    retained: content(resources).map((v) => ({
      kind: v.kind,
      resourceId: v.resourceId,
      ownership: v.ownership,
      fileCount: v.fileCount,
      bytes: v.bytes,
    })),
  })
  async function execute(
    s: PresentationLifecycleScope,
    rev: number,
    id: string,
    signal?: AbortSignal,
  ) {
    check(signal)
    current(s, rev, id)
    // Drain outside the project lock: a worker may need that lock to settle.
    await stopPresentationProjectWork({ root: options.userDataPath, ...s })
    check(signal)
    current(s, rev, id)
    const release = await options.acquireProjectLock(s.projectId)
    try {
      check(signal)
      let r = current(s, rev, id)
      await inventory(s, signal)
      current(s, rev, id)
      const entries = [...r.deletion!.resources].sort(
        (a, b) => Number(a.kind === 'project') - Number(b.kind === 'project'),
      )
      for (const entry of entries) {
        check(signal)
        current(s, rev, id)
        if (entry.status === 'removed' || entry.status === 'reference_removed') continue
        if (entry.kind === 'project') {
          const fresh = await inventory(s, signal)
          current(s, rev, id)
          if (content(fresh.resources).some((v) => v.kind !== 'project')) continue
          if (
            r.deletion!.resources.some(
              (v) => v.kind !== 'project' && !['removed', 'reference_removed'].includes(v.status),
            )
          )
            continue
        }
        const guard = () => {
          check(signal)
          current(s, rev, id)
        }
        const result = await removePresentationProjectDeletionResource({
          userDataPath: options.userDataPath,
          scope: s,
          deletionId: id,
          resourceId: entry.resourceId,
          expectedRevision: rev,
          assertDeleting: guard,
        })
        guard()
        r = life.recordDeletionResult(s, rev, result)
        rev = r.revision
      }
      const fresh = await inventory(s, signal)
      check(signal)
      r = current(s, rev, id)
      if (
        content(fresh.resources).length === 0 &&
        r.deletion!.resources.every((v) => ['removed', 'reference_removed'].includes(v.status))
      )
        r = life.finishDeletion(s, rev, id)
      return report(r, fresh.resources)
    } finally {
      release()
    }
  }
  const service = {
    async preview(input: PresentationLifecycleScope, signal?: AbortSignal) {
      const s = scope(input)
      check(signal)
      if (!store.projectScope(s.projectId, s.documentId)) throw Error('project_not_found')
      const r = life.read(s)
      if (r && r.state !== 'active') throw Error('project_' + r.state)
      const i = await inventory(s, signal)
      check(signal)
      const now = life.read(s)
      if (now?.revision !== r?.revision || now?.state !== r?.state) throw Error('revision_conflict')
      if (!store.projectScope(s.projectId, s.documentId)) throw Error('project_not_found')
      return {
        expectedRevision: r?.revision ?? null,
        confirmationToken: token(s, r?.revision ?? null, i.resources),
        resources: content(i.resources),
        governanceRetained: true as const,
      }
    },
    async confirm(
      input: {
        scope: PresentationLifecycleScope
        expectedRevision: number | null
        confirmationToken: string
        deletionId: string
      },
      signal?: AbortSignal,
    ) {
      const body = owned(input, ['scope', 'expectedRevision', 'confirmationToken', 'deletionId']),
        s = scope(body.scope)
      revision(body.expectedRevision, true)
      if (
        typeof body.confirmationToken !== 'string' ||
        !/^[a-f0-9]{64}$/.test(body.confirmationToken) ||
        typeof body.deletionId !== 'string' ||
        !/^[A-Za-z0-9_-]{1,128}$/.test(body.deletionId)
      )
        throw Error('invalid_request')
      check(signal)
      const release = await options.acquireProjectLock(s.projectId)
      let r: PresentationLifecycleRecord
      try {
        check(signal)
        const old = life.read(s)
        if ((old?.revision ?? null) !== body.expectedRevision) throw Error('revision_conflict')
        if (old && old.state !== 'active') throw Error('project_' + old.state)
        if (!store.projectScope(s.projectId, s.documentId)) throw Error('project_not_found')
        const i = await inventory(s, signal)
        check(signal)
        const after = life.read(s)
        if ((after?.revision ?? null) !== body.expectedRevision || after?.state !== old?.state)
          throw Error('revision_conflict')
        if (token(s, body.expectedRevision, i.resources) !== body.confirmationToken)
          throw Error('confirmation_conflict')
        const resources = content(i.resources)
          .filter(known)
          .map((v) => ({
            resourceId: v.resourceId,
            kind: v.kind,
            ownership:
              v.ownership === 'project_exclusive'
                ? ('project_exclusive' as const)
                : v.ownership === 'unproven'
                  ? ('unproven' as const)
                  : ('shared_reference' as const),
          }))
        if (!resources.length) throw Error('deletion_incomplete')
        if (!store.projectScope(s.projectId, s.documentId)) throw Error('project_not_found')
        r = old ?? life.initialize(s)
        r = life.beginDeletion(s, r.revision, {
          deletionId: body.deletionId,
          reason: 'user',
          resources,
        })
      } finally {
        release()
      }
      return execute(s, r.revision, body.deletionId, signal)
    },
    async resume(
      input: { scope: PresentationLifecycleScope; expectedRevision: number; deletionId: string },
      signal?: AbortSignal,
    ) {
      const body = owned(input, ['scope', 'expectedRevision', 'deletionId']),
        s = scope(body.scope)
      revision(body.expectedRevision)
      if (typeof body.deletionId !== 'string' || !/^[A-Za-z0-9_-]{1,128}$/.test(body.deletionId))
        throw Error('invalid_request')
      return execute(s, body.expectedRevision, body.deletionId, signal)
    },
  }
  const attempt = async <T>(work: () => Promise<T>): Promise<T> => {
    try {
      return await work()
    } catch (error) {
      let message = ''
      try {
        message = error instanceof Error ? error.message : ''
      } catch {
        /* No raw transport errors. */
      }
      // eslint-disable-next-line preserve-caught-error -- Public errors must not retain private transport causes.
      throw Error(
        [
          'invalid_request',
          'invalid_state',
          'aborted',
          'cancelled',
          'project_not_found',
          'project_deleting',
          'project_deleted',
          'document_mismatch',
          'revision_conflict',
          'deletion_conflict',
          'confirmation_conflict',
          'deletion_incomplete',
          'output_too_large',
          'presentation_inventory_invalid',
          'presentation_inventory_budget',
        ].includes(message)
          ? message
          : 'invalid_state',
      )
    }
  }
  return {
    preview: (...args: Parameters<typeof service.preview>) =>
      attempt(() => service.preview(...args)),
    confirm: (...args: Parameters<typeof service.confirm>) =>
      attempt(() => service.confirm(...args)),
    resume: (...args: Parameters<typeof service.resume>) => attempt(() => service.resume(...args)),
  }
}
