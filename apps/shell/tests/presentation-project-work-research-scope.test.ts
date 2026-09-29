import { expect, it } from 'vitest'
import {
  registerPresentationProjectWork,
  stopPresentationProjectWork,
} from '../src/main/presentation-project-work'
it('drains the exact legacy research document scope up to 4096 characters', async () => {
  const scope = {
    root: '/tmp/project-work-research-scope',
    projectId: 'research',
    documentId: 'd'.repeat(4096),
  }
  const work = registerPresentationProjectWork({ scope })
  try {
    const drain = stopPresentationProjectWork(scope)
    expect(work.signal.aborted).toBe(true)
    work.finish()
    await drain
    expect(() =>
      registerPresentationProjectWork({ scope: { ...scope, documentId: 'd'.repeat(4097) } }),
    ).toThrow('invalid_request')
  } finally {
    work.finish()
  }
})
