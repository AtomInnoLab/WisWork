import { PresentationLifecycleStore, type PresentationLifecycleScope } from '@wiswork/project-store'

export interface PresentationProjectWriteLease {
  readonly scope: Readonly<PresentationLifecycleScope>
  readonly revision: number
  assertWritable(): void
}

/** A request owns one lifecycle revision; this guard never renews it or creates project content. */
export function capturePresentationProjectWriteLease(options: {
  store: PresentationLifecycleStore
  scope: PresentationLifecycleScope
  readExistingProject: (
    scope: Readonly<PresentationLifecycleScope>,
  ) => PresentationLifecycleScope | undefined
  signal?: AbortSignal
}): PresentationProjectWriteLease {
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
    record = store.initialize(scope)
  }
  check()
  const revision = record.revision
  store.assertActive(scope, revision)
  return Object.freeze({
    scope,
    revision,
    assertWritable() {
      check()
      store.assertActive(scope, revision)
      check()
    },
  })
}
