import {
  MAX_PRESENTATION_GOVERNANCE_RESPONSE_BYTES,
  parsePresentationGovernancePreview,
  parsePresentationGovernancePolicy,
  parsePresentationGovernanceDeletionReport,
  parsePresentationGovernanceLifecycle,
  parsePresentationGovernanceAudit,
  parsePresentationProjectDeletionAttempt,
  type PresentationGovernancePreview,
  type PresentationGovernanceDeletionReport,
  type PresentationLifecycleRecord,
  type PresentationGovernanceAudit,
  type PresentationLifecycleScope,
  type PresentationProjectDeletionAttempt,
  type PresentationRetentionPolicy,
} from '../../../../packages/project-store/src/presentation-project-governance'
export interface PresentationGovernanceSnapshot {
  available: boolean
  phase: 'idle' | 'busy' | 'unknown' | 'partial' | 'deleted' | 'error'
  scope?: PresentationLifecycleScope
  preview?: PresentationGovernancePreview
  lifecycle?: PresentationLifecycleRecord | null
  deletion?: PresentationGovernanceDeletionReport
  attempt?: PresentationProjectDeletionAttempt
  audit?: PresentationGovernanceAudit
  error?: string
}
export interface PresentationGovernanceOptions {
  available(): boolean
  documentId(): Promise<string>
  currentProjectId(): string | undefined
  request(body: Record<string, unknown>, signal?: AbortSignal): Promise<Response>
  readAttempt(scope: PresentationLifecycleScope): unknown
  writeAttempt(
    scope: PresentationLifecycleScope,
    attempt: PresentationProjectDeletionAttempt | undefined,
  ): void
}
const canonical = (value: unknown) =>
  JSON.stringify(value, (_key, v: unknown) =>
    v && typeof v === 'object' && !Array.isArray(v)
      ? Object.fromEntries(Object.entries(v).sort(([a], [b]) => a.localeCompare(b)))
      : v,
  )
export function createPresentationProjectGovernanceController(
  options: PresentationGovernanceOptions,
) {
  let state: PresentationGovernanceSnapshot = { available: options.available(), phase: 'idle' },
    epoch = 0,
    active: AbortController | undefined
  const listeners = new Set<() => void>()
  const publish = (value: PresentationGovernanceSnapshot) => {
    const cloned = structuredClone(value)
    const freeze = (v: unknown): void => {
      if (v && typeof v === 'object') {
        Object.values(v).forEach(freeze)
        Object.freeze(v)
      }
    }
    freeze(cloned)
    state = cloned
    listeners.forEach((f) => f())
  }
  const clear = () => {
    epoch++
    active?.abort()
    active = undefined
    publish({ available: options.available(), phase: 'idle' })
  }
  async function run(
    action: (
      scope: PresentationLifecycleScope,
      call: (operation: string, key: string, fields?: Record<string, unknown>) => Promise<unknown>,
      check: () => Promise<void>,
    ) => Promise<void>,
    mutation = false,
  ) {
    if (!options.available() || state.phase === 'busy') {
      if (!options.available()) clear()
      return
    }
    active?.abort()
    const controller = new AbortController()
    active = controller
    const generation = ++epoch
    let scope: PresentationLifecycleScope | undefined
    const checkSync = () => {
      if (
        controller.signal.aborted ||
        generation !== epoch ||
        !options.available() ||
        options.currentProjectId() !== scope?.projectId
      )
        throw Error('stale')
    }
    const check = async () => {
      checkSync()
      const doc = await options.documentId()
      checkSync()
      if (doc !== scope?.documentId) throw Error('stale')
    }
    try {
      const projectId = options.currentProjectId()
      if (!projectId) return
      const documentId = await options.documentId()
      if (
        controller.signal.aborted ||
        generation !== epoch ||
        !options.available() ||
        options.currentProjectId() !== projectId
      )
        throw Error('stale')
      scope = { documentId, projectId }
      parsePresentationProjectDeletionAttempt({
        version: 1,
        scope,
        expectedRevision: null,
        confirmationToken: '0'.repeat(64),
        deletionId: 'scope',
      })
      if (
        mutation &&
        (!state.scope ||
          state.scope.documentId !== documentId ||
          state.scope.projectId !== projectId)
      )
        throw Error('stale')
      if (
        state.scope &&
        (state.scope.documentId !== documentId || state.scope.projectId !== projectId)
      )
        publish({ available: true, phase: 'idle' })
      const previous = state
      publish({ ...previous, scope, available: true, phase: 'busy', error: undefined })
      const call = async (operation: string, key: string, fields: Record<string, unknown> = {}) => {
        await check()
        checkSync()
        const response = await options.request(
          { operation, ...scope, ...structuredClone(fields) },
          controller.signal,
        )
        await check()
        if (!response.ok) throw Error('unavailable')
        const text = await response.text()
        await check()
        if (new TextEncoder().encode(text).byteLength > MAX_PRESENTATION_GOVERNANCE_RESPONSE_BYTES)
          throw Error('invalid')
        const value: unknown = JSON.parse(text)
        if (
          !value ||
          typeof value !== 'object' ||
          Array.isArray(value) ||
          Object.keys(value).length !== 1 ||
          !Object.hasOwn(value, key)
        )
          throw Error('invalid')
        return (value as Record<string, unknown>)[key]
      }
      await action(scope, call, check)
      await check()
      if ((state as PresentationGovernanceSnapshot).phase === 'busy')
        publish({ ...state, phase: 'idle' })
    } catch (error) {
      if (generation === epoch) {
        if (error instanceof Error && error.message === 'stale') {
          publish({ available: options.available(), phase: 'idle' })
          return
        }
        publish({
          ...state,
          available: options.available(),
          phase: mutation ? 'unknown' : 'error',
          preview: undefined,
          audit: undefined,
          error: mutation ? '操作结果尚未确认，请核对原请求状态。' : '本机项目治理暂不可用。',
        })
      }
    } finally {
      if (active === controller) active = undefined
    }
  }
  function boundLifecycle(value: unknown, scope: PresentationLifecycleScope) {
    const previous = state.lifecycle
    if (value === null) {
      if (previous) throw Error('invalid')
      return null
    }
    const record = parsePresentationGovernanceLifecycle(value)
    if (record.projectId !== scope.projectId || record.documentId !== scope.documentId)
      throw Error('invalid')
    if (
      previous &&
      previous.documentId === scope.documentId &&
      previous.projectId === scope.projectId
    ) {
      if (
        record.revision < previous.revision ||
        record.anonymousProjectId !== previous.anonymousProjectId ||
        record.createdAt !== previous.createdAt ||
        canonical(record.audit.slice(0, previous.audit.length)) !== canonical(previous.audit) ||
        (record.revision === previous.revision && canonical(record) !== canonical(previous))
      )
        throw Error('invalid')
    }
    return record
  }
  return {
    snapshot: () => state,
    subscribe: (listener: () => void) => {
      listeners.add(listener)
      return () => {
        listeners.delete(listener)
      }
    },
    clear,
    cancel: clear,
    refresh: () =>
      run(async (scope, call) => {
        const lifecycle = boundLifecycle(await call('project_lifecycle_read', 'lifecycle'), scope)
        const raw = options.readAttempt(scope)
        const attempt = raw === undefined ? undefined : parsePresentationProjectDeletionAttempt(raw)
        if (
          attempt &&
          (attempt.scope.projectId !== scope.projectId ||
            attempt.scope.documentId !== scope.documentId)
        )
          throw Error('invalid')
        publish({
          ...state,
          lifecycle,
          attempt,
          phase: attempt ? 'unknown' : 'idle',
          preview: undefined,
        })
      }),
    preview: () =>
      run(async (scope, call) => {
        const saved = options.readAttempt(scope)
        if (saved !== undefined) {
          const attempt = parsePresentationProjectDeletionAttempt(saved)
          if (
            attempt.scope.documentId !== scope.documentId ||
            attempt.scope.projectId !== scope.projectId
          )
            throw Error('invalid')
          publish({ ...state, attempt, phase: 'unknown', preview: undefined })
          return
        }
        if (state.attempt) throw Error('pending')
        const preview = parsePresentationGovernancePreview(
          await call('project_deletion_preview', 'preview'),
        )
        publish({ ...state, preview })
      }),
    confirmDeletion: () => {
      const original = state.preview
      if (!original || state.attempt) return Promise.resolve()
      return run(async (scope, call, check) => {
        const attempt = parsePresentationProjectDeletionAttempt({
          version: 1,
          scope,
          expectedRevision: original.expectedRevision,
          confirmationToken: original.confirmationToken,
          deletionId: crypto.randomUUID(),
        })
        await check()
        options.writeAttempt(scope, structuredClone(attempt))
        publish({ ...state, attempt, preview: undefined })
        const deletion = parsePresentationGovernanceDeletionReport(
          await call('project_deletion_confirm', 'deletion', {
            expectedRevision: attempt.expectedRevision,
            confirmationToken: attempt.confirmationToken,
            deletionId: attempt.deletionId,
          }),
        )
        if (deletion.deletionId !== attempt.deletionId) throw Error('invalid')
        publish({ ...state, deletion, phase: deletion.state })
      }, true)
    },
    checkAttempt: () =>
      run(async (scope, call) => {
        const raw = state.attempt ?? options.readAttempt(scope)
        if (!raw) throw Error('missing')
        const attempt = parsePresentationProjectDeletionAttempt(raw)
        if (
          attempt.scope.documentId !== scope.documentId ||
          attempt.scope.projectId !== scope.projectId
        )
          throw Error('invalid')
        const lifecycle = boundLifecycle(await call('project_lifecycle_read', 'lifecycle'), scope)
        if (!lifecycle || lifecycle.deletion?.deletionId !== attempt.deletionId) {
          publish({ ...state, attempt, lifecycle, phase: 'unknown' })
          return
        }
        publish({
          ...state,
          attempt,
          lifecycle,
          phase: lifecycle.state === 'deleted' ? 'deleted' : 'partial',
        })
      }),
    resumeDeletion: () => {
      const attempt = state.attempt,
        lifecycle = state.lifecycle
      if (
        !attempt ||
        !lifecycle ||
        lifecycle.state !== 'deleting' ||
        lifecycle.deletion?.deletionId !== attempt.deletionId
      )
        return Promise.resolve()
      return run(async (_scope, call) => {
        const deletion = parsePresentationGovernanceDeletionReport(
          await call('project_deletion_resume', 'deletion', {
            expectedRevision: lifecycle.revision,
            deletionId: attempt.deletionId,
          }),
        )
        if (deletion.deletionId !== attempt.deletionId) throw Error('invalid')
        publish({ ...state, deletion, phase: deletion.state })
      }, true)
    },
    initializePolicy: () => {
      if (!state.preview || state.attempt || state.lifecycle) return Promise.resolve()
      return run(async (scope, call) => {
        const lifecycle = boundLifecycle(
          await call('project_lifecycle_initialize', 'lifecycle'),
          scope,
        )
        if (!lifecycle || lifecycle.state !== 'active') throw Error('invalid')
        publish({ ...state, lifecycle, preview: undefined })
      }, true)
    },
    setPolicy: (policy: PresentationRetentionPolicy) => {
      const lifecycle = state.lifecycle
      if (!lifecycle || lifecycle.state !== 'active') return Promise.resolve()
      let frozen: PresentationRetentionPolicy
      try {
        frozen = parsePresentationGovernancePolicy(policy)
      } catch {
        publish({ ...state, phase: 'error', error: '保留天数须为 1 至 36500 的整数，或留空。' })
        return Promise.resolve()
      }
      return run(async (scope, call) => {
        const next = boundLifecycle(
          await call('project_lifecycle_set_policy', 'lifecycle', {
            expectedRevision: lifecycle.revision,
            policy: frozen,
          }),
          scope,
        )
        if (
          !next ||
          next.state !== 'active' ||
          next.revision !== lifecycle.revision + 1 ||
          canonical(next.policy) !== canonical(frozen)
        )
          throw Error('invalid')
        publish({ ...state, lifecycle: next, preview: undefined })
      }, true)
    },
    exportAudit: () =>
      run(async (_scope, call) => {
        const audit = parsePresentationGovernanceAudit(
          await call('project_lifecycle_export_audit', 'audit'),
        )
        if (
          state.lifecycle &&
          (audit.anonymousProjectId !== state.lifecycle.anonymousProjectId ||
            canonical(audit.events.slice(0, state.lifecycle.audit.length)) !==
              canonical(state.lifecycle.audit))
        )
          throw Error('invalid')
        publish({ ...state, audit })
      }),
  }
}
