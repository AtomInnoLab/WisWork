import {
  parsePresentationProjectDeletionAttempt,
  type PresentationLifecycleScope,
  type PresentationProjectDeletionAttempt,
} from '@wiswork/project-store/presentation-project-governance'
export function createPresentationGovernanceStorage(
  storage: Pick<Storage, 'getItem' | 'setItem' | 'removeItem'>,
) {
  const key = (scope: PresentationLifecycleScope) =>
    'wiswork.governance.intent.v1:' + JSON.stringify([scope.documentId, scope.projectId])
  return {
    read(scope: PresentationLifecycleScope): PresentationProjectDeletionAttempt | undefined {
      const raw = storage.getItem(key(scope))
      if (raw === null) return
      const attempt = parsePresentationProjectDeletionAttempt(JSON.parse(raw))
      if (
        attempt.scope.documentId !== scope.documentId ||
        attempt.scope.projectId !== scope.projectId
      )
        throw Error('presentation_governance_response_invalid')
      return attempt
    },
    write(
      scope: PresentationLifecycleScope,
      value: PresentationProjectDeletionAttempt | undefined,
    ) {
      if (value === undefined) {
        storage.removeItem(key(scope))
        return
      }
      const attempt = parsePresentationProjectDeletionAttempt(value)
      if (
        attempt.scope.documentId !== scope.documentId ||
        attempt.scope.projectId !== scope.projectId
      )
        throw Error('presentation_governance_response_invalid')
      storage.setItem(key(scope), JSON.stringify(attempt))
    },
  }
}
