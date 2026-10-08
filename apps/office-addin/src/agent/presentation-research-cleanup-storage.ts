import type { PresentationResearchDeleteAttempt } from './presentation-research.js'
/** Only finite cleanup identity metadata is persisted; original research and files stay on PC. */
export function createPresentationResearchDeletePersistence(
  documentId: string,
  storage?: Pick<Storage, 'getItem' | 'setItem' | 'removeItem'>,
) {
  const key = `wiswork:research-delete-attempt:${encodeURIComponent(documentId)}`
  const check = (scope: string) => {
    if (scope !== documentId) throw new Error('presentation_document_changed')
  }
  return {
    read(document: string): unknown {
      check(document)
      const text = storage?.getItem(key)
      if (!text) return undefined
      if (new TextEncoder().encode(text).length > 4096)
        throw new Error('presentation_response_invalid')
      return JSON.parse(text)
    },
    write(document: string, attempt: PresentationResearchDeleteAttempt | undefined) {
      check(document)
      if (!storage) throw new Error('presentation_cleanup_recovery_unavailable')
      if (attempt && attempt.documentId !== documentId)
        throw new Error('presentation_document_changed')
      if (attempt) storage.setItem(key, JSON.stringify(attempt))
      else storage.removeItem(key)
    },
  }
}
