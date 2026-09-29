import { PresentationLifecycleStore, type PresentationLifecycleScope } from '@wiswork/project-store'

interface LeaseOptions {
  store: PresentationLifecycleStore
  scope: PresentationLifecycleScope
  readExistingProject: (
    scope: Readonly<PresentationLifecycleScope>,
  ) => PresentationLifecycleScope | undefined
  signal?: AbortSignal
}
export interface PresentationProjectWriteLease {
  readonly scope: Readonly<PresentationLifecycleScope>
  readonly revision: number
  assertWritable(): void
}
export interface PresentationProjectReadLease {
  readonly scope: Readonly<PresentationLifecycleScope>
  readonly revision: number | undefined
  assertCurrent(): void
}

function capture(options: LeaseOptions, initializeMissing: boolean): PresentationProjectReadLease {
  const { store, signal, readExistingProject } = options
  const scope = Object.freeze({
    projectId: options.scope.projectId,
    documentId: options.scope.documentId,
  })
  const check = () => {
    if (signal?.aborted) throw Error('aborted')
  }
  check()
  let record = store.read(scope)
  if (!record) {
    const existing = readExistingProject(scope)
    check()
    if (
      !existing ||
      existing.projectId !== scope.projectId ||
      existing.documentId !== scope.documentId
    )
      throw Error('project_not_found')
    if (initializeMissing) record = store.initialize(scope)
  }
  const revision = record?.revision
  const assertCurrent = () => {
    check()
    if (revision === undefined) {
      if (store.read(scope) !== undefined) throw Error('revision_conflict')
    } else store.assertActive(scope, revision)
    check()
  }
  assertCurrent()
  return Object.freeze({ scope, revision, assertCurrent })
}

/** A write request owns one lifecycle revision; this guard never renews it or creates project content. */
export function capturePresentationProjectWriteLease(
  options: LeaseOptions,
): PresentationProjectWriteLease {
  const lease = capture(options, true)
  return Object.freeze({
    scope: lease.scope,
    revision: lease.revision!,
    assertWritable: lease.assertCurrent,
  })
}

/** Pure reads preserve absent legacy control metadata and grant no write permission. */
export function capturePresentationProjectReadLease(
  options: LeaseOptions,
): PresentationProjectReadLease {
  return capture(options, false)
}

/** Only internally validated creation entrypoints may use this before their first content write. */
export function capturePresentationProjectCreationLease(
  options: Omit<LeaseOptions, 'readExistingProject'>,
): PresentationProjectWriteLease {
  const scope = { projectId: options.scope.projectId, documentId: options.scope.documentId }
  return capturePresentationProjectWriteLease({
    store: options.store,
    scope,
    signal: options.signal,
    readExistingProject: () => scope,
  })
}

/** Async ownership proof cannot refresh the lifecycle epoch captured before its first await. */
export async function capturePresentationProjectAsyncReadLease(
  options: Omit<LeaseOptions, 'readExistingProject'> & {
    readExistingProject: (
      scope: Readonly<PresentationLifecycleScope>,
    ) => Promise<PresentationLifecycleScope | undefined>
  },
): Promise<PresentationProjectReadLease> {
  const { store, signal, readExistingProject } = options
  const scope = Object.freeze({
    projectId: options.scope.projectId,
    documentId: options.scope.documentId,
  })
  const check = () => {
    if (signal?.aborted) throw Error('aborted')
  }
  check()
  const revision = store.read(scope)?.revision
  const assertCurrent = () => {
    check()
    if (revision === undefined) {
      if (store.read(scope) !== undefined) throw Error('revision_conflict')
    } else store.assertActive(scope, revision)
    check()
  }
  assertCurrent()
  const existing = await readExistingProject(scope)
  assertCurrent()
  if (
    !existing ||
    existing.projectId !== scope.projectId ||
    existing.documentId !== scope.documentId
  )
    throw Error('project_not_found')
  return Object.freeze({ scope, revision, assertCurrent })
}
