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
