import type { PresentationResearchAbandonAttempt } from './presentation-research.js'
/** Only finite recovery identity metadata is persisted; original research and files stay on PC. */
export function createPresentationResearchAbandonPersistence(
  documentId: string,
  storage?: Pick<Storage, 'getItem' | 'setItem' | 'removeItem'>,
) {
  const key = `wiswork:research-abandon-attempt:${encodeURIComponent(documentId)}`
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
    write(document: string, attempt: PresentationResearchAbandonAttempt | undefined) {
      check(document)
      if (!storage) throw new Error('presentation_abandon_recovery_unavailable')
      if (attempt && attempt.documentId !== documentId)
        throw new Error('presentation_document_changed')
      if (attempt) {
        if (
          Object.keys(attempt).sort().join(',') !==
          'documentId,draftDigest,expectedRevision,ledgerId,projectId,sequence'
        )
          throw new Error('presentation_response_invalid')
        const text = JSON.stringify(attempt)
        if (new TextEncoder().encode(text).length > 4096)
          throw new Error('presentation_response_invalid')
        storage.setItem(key, text)
      } else storage.removeItem(key)
    },
  }
}
