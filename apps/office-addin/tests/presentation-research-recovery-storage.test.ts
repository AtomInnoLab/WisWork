import { expect, it, vi } from 'vitest'
import { createPresentationResearchAbandonPersistence } from '../src/agent/presentation-research-recovery-storage.js'
it('persists only finite document-scoped recovery identity and recovers on taskpane reopening', () => {
  const values = new Map<string, string>()
  const storage = {
    getItem: (key: string) => values.get(key) ?? null,
    setItem: (key: string, value: string) => {
      values.set(key, value)
    },
    removeItem: vi.fn((key: string) => {
      values.delete(key)
    }),
  }
  const attempt = {
    documentId: 'doc',
    projectId: 'research',
    ledgerId: 'ledger1',
    sequence: 1,
    draftDigest: 'a'.repeat(64),
    expectedRevision: 1,
  }
  const persistence = createPresentationResearchAbandonPersistence('doc', storage)
  persistence.write('doc', attempt)
  expect(createPresentationResearchAbandonPersistence('doc', storage).read('doc')).toEqual(attempt)
  expect(
    createPresentationResearchAbandonPersistence('other', storage).read('other'),
  ).toBeUndefined()
  expect(() => persistence.write('other', attempt)).toThrow('presentation_document_changed')
  expect(() => persistence.write('doc', { ...attempt, draft: 'secret' } as typeof attempt)).toThrow(
    'presentation_response_invalid',
  )
  persistence.write('doc', undefined)
  expect(persistence.read('doc')).toBeUndefined()
  expect(() => createPresentationResearchAbandonPersistence('doc').write('doc', attempt)).toThrow(
    'presentation_abandon_recovery_unavailable',
  )
})
