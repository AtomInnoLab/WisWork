import { readFile } from 'node:fs/promises'
import { join } from 'node:path'
import { describe, expect, it } from 'vitest'
import { fingerprintPresentation, openPptx } from '@wiswork/pptx-engine'
import {
  inspectSlidesAcceptanceAuthority,
  inspectSlidesAcceptanceLease,
} from '../src/main/operations/acceptance-authority'
import type { Session } from '../src/main/session-state'

describe('PC authoritative acceptance inspection', () => {
  it('publishes revision-bound durable identities and locked decoration facts without mutation', async () => {
    const bytes = await readFile(
      join(__dirname, '../../../packages/pptx-engine/tests/fixtures/01_standard_business.pptx'),
    )
    const opened = await openPptx(bytes)
    const slide = opened.deck.slides[0]!
    const element = slide.elements[0]!
    const emphasis = slide.elements[1]!
    element.creationId = '{00000000-0000-4000-8000-000000000111}'
    element.placeholder = 'title'
    emphasis.creationId = '{00000000-0000-4000-8000-000000000333}'
    emphasis.placeholder = 'subTitle'
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
    const lease = await inspectSlidesAcceptanceLease(session)
    expect(Object.keys(lease!)).toEqual([
      'documentToken',
      'sessionToken',
      'revision',
      'slideCount',
      'leaseToken',
    ])
    const revision = await fingerprintPresentation(opened)
    const request = {
      affectedSlides: [1],
      referenceSlides: [],
      expectedDocumentToken: 'document-production',
      expectedSessionToken: 'session-production',
      expectedRevision: revision,
      leaseToken: lease!.leaseToken,
    }
    const snapshot = await inspectSlidesAcceptanceAuthority(session, request)
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
        expect.objectContaining({ role: 'emphasis', locked: false }),
        expect.objectContaining({
          role: 'title',
          locked: true,
          targetToken: expect.stringMatching(/^target:[0-9a-f]{64}$/),
        }),
      ]),
    )
    expect(snapshot!.slides).toHaveLength(1)
    expect(snapshot!.slides[0]!.elements.every((fact) => !('text' in fact.properties))).toBe(true)
    expect(
      await inspectSlidesAcceptanceAuthority(session, {
        ...request,
        expectedRevision: `sha256:${'f'.repeat(64)}`,
      }),
    ).toBeNull()
    session.sessionInstanceId = 'session-replaced'
    expect(await inspectSlidesAcceptanceAuthority(session, request)).toBeNull()
    expect(
      JSON.stringify({
        creationId: element.creationId,
        decorationId: slide.decorations[0]!.creationId,
        dirty: element.dirty,
        generation: session.mutationGeneration,
      }),
    ).toBe(before)
  })

  it('inspects only the bounded requested union in decks larger than fifty slides', async () => {
    const bytes = await readFile(
      join(__dirname, '../../../packages/pptx-engine/tests/fixtures/01_standard_business.pptx'),
    )
    const opened = await openPptx(bytes)
    opened.deck.slides = Array.from({ length: 60 }, (_, index) => ({
      ...opened.deck.slides[0]!,
      durableId: `ppt/slides/slide${index + 1}.xml`,
      elements: [],
    }))
    const session = {
      opened,
      documentInstanceId: 'doc-large',
      sessionInstanceId: 'session-large',
    } as Session
    const revision = await fingerprintPresentation(opened)
    const lease = await inspectSlidesAcceptanceLease(session)
    const snapshot = await inspectSlidesAcceptanceAuthority(session, {
      affectedSlides: [1, 60],
      referenceSlides: [30],
      expectedDocumentToken: 'doc-large',
      expectedSessionToken: 'session-large',
      expectedRevision: revision,
      leaseToken: lease!.leaseToken,
    })
    expect(snapshot!.slides.map((slide) => slide.number)).toEqual([1, 30, 60])
    await expect(
      inspectSlidesAcceptanceAuthority(session, {
        affectedSlides: Array.from({ length: 51 }, (_, index) => index + 1),
        referenceSlides: [],
        expectedDocumentToken: 'doc-large',
        expectedSessionToken: 'session-large',
        expectedRevision: revision,
        leaseToken: lease!.leaseToken,
      }),
    ).rejects.toThrow(/bounded/)
  })

  it('returns contract-bound text digests without exposing raw text', async () => {
    const bytes = await readFile(
      join(__dirname, '../../../packages/pptx-engine/tests/fixtures/01_standard_business.pptx'),
    )
    const opened = await openPptx(bytes)
    const element = opened.deck.slides[0]!.elements[0]!
    element.creationId = '{00000000-0000-4000-8000-000000000444}'
    element.placeholder = 'title'
    const session = {
      opened,
      documentInstanceId: 'doc-text',
      sessionInstanceId: 'session-text',
    } as Session
    const lease = await inspectSlidesAcceptanceLease(session)
    const base = {
      affectedSlides: [1],
      referenceSlides: [],
      expectedDocumentToken: lease!.documentToken,
      expectedSessionToken: lease!.sessionToken,
      expectedRevision: lease!.revision,
      leaseToken: lease!.leaseToken,
    }
    const structural = await inspectSlidesAcceptanceAuthority(session, base)
    const targetToken = structural!.slides[0]!.elements.find(
      (fact) => fact.role === 'title',
    )!.targetToken
    const verified = await inspectSlidesAcceptanceAuthority(session, {
      ...base,
      textChecks: [{ checkId: 'check-text', targetToken, expectedText: 'Q3 Business Review' }],
    })
    const fact = verified!.slides[0]!.elements.find((item) => item.targetToken === targetToken)!
    expect(verified!.textMatches?.['check-text']).toMatchObject({ targetToken, matches: true })
    expect(verified!.textMatches?.['check-text']?.proof).toMatch(/^sha256:[0-9a-f]{64}$/)
    expect(JSON.stringify(verified)).not.toContain('Q3 Business Review')
    expect(fact.properties.text).toBeUndefined()
    const changed = await inspectSlidesAcceptanceAuthority(session, {
      ...base,
      textChecks: [{ checkId: 'check-text', targetToken, expectedText: 'changed' }],
    })
    expect(changed!.textMatches?.['check-text']?.matches).toBe(false)
  })
})
