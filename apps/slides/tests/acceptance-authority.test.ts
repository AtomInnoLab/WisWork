import { readFile } from 'node:fs/promises'
import { join } from 'node:path'
import { describe, expect, it } from 'vitest'
import { openPptx } from '@wiswork/pptx-engine'
import { inspectSlidesAcceptanceAuthority } from '../src/main/operations/acceptance-authority'
import type { Session } from '../src/main/session-state'

describe('PC authoritative acceptance inspection', () => {
  it('publishes revision-bound durable identities and locked decoration facts without mutation', async () => {
    const bytes = await readFile(
      join(__dirname, '../../../packages/pptx-engine/tests/fixtures/01_standard_business.pptx'),
    )
    const opened = await openPptx(bytes)
    const slide = opened.deck.slides[0]!
    const element = slide.elements[0]!
    element.creationId = '{00000000-0000-4000-8000-000000000111}'
    element.placeholder = 'title'
    slide.decorations = [{ ...element, creationId: '{00000000-0000-4000-8000-000000000222}' }]
    const session = {
      opened,
      documentInstanceId: 'document-production',
      sessionInstanceId: 'session-production',
    } as Session
    const before = JSON.stringify({
      creationId: element.creationId,
      decorationId: slide.decorations[0]!.creationId,
      dirty: element.dirty,
      generation: session.mutationGeneration,
    })
    const snapshot = await inspectSlidesAcceptanceAuthority(session)
    expect(snapshot).toMatchObject({
      documentToken: 'document-production',
      sessionToken: 'session-production',
    })
    expect(snapshot!.revision).toMatch(/^sha256:[0-9a-f]{64}$/)
    expect(snapshot!.slides[0]!.elements).toEqual(
      expect.arrayContaining([
        expect.objectContaining({
          role: 'title',
          locked: false,
          targetToken: expect.stringMatching(/^target:[0-9a-f]{64}$/),
        }),
        expect.objectContaining({
          role: 'title',
          locked: true,
          targetToken: expect.stringMatching(/^target:[0-9a-f]{64}$/),
        }),
      ]),
    )
    expect(
      JSON.stringify({
        creationId: element.creationId,
        decorationId: slide.decorations[0]!.creationId,
        dirty: element.dirty,
        generation: session.mutationGeneration,
      }),
    ).toBe(before)
  })
})
