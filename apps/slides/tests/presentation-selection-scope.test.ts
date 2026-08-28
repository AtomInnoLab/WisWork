import { describe, expect, it, vi } from 'vitest'
import { executePreparedGeometryFamilyTransaction } from '../src/renderer/ai/presentation-geometry-transactions'
import type { SelectionScope } from '../src/renderer/ai/edit-queue'
import { createSlidesSkill, type DeckAccess } from '../src/renderer/ai/slides-skill'

const fingerprint = `sha256:${'a'.repeat(64)}`
const scope: SelectionScope = {
  documentId: 'doc',
  sessionId: 'session',
  generation: 1,
  slides: [
    {
      slideId: 'slide-1',
      elements: [
        { elementId: 'selected', expectedType: 'shape', expectedFingerprint: fingerprint },
      ],
    },
  ],
}

describe('canonical transaction selection scope', () => {
  it('does not advertise legacy mutation tools during a scoped run', () => {
    const skill = createSlidesSkill({
      getSlides: () => [],
      getCurrent: () => 0,
      getSelectedIds: () => [],
      getSelectionScope: () => scope,
      applySlide: vi.fn(),
      applyDeck: vi.fn(),
      fitWidthPx: 1280,
    } satisfies DeckAccess)
    const names = skill.tools.map((tool) => tool.name)
    expect(names).toContain('set_element_transform')
    expect(names).not.toContain('add_slide')
    expect(names).not.toContain('insert_web_image')
  })

  it('rejects an out-of-scope prepared target before transaction dispatch', async () => {
    const executePresentationTransaction = vi.fn()
    const cancelPresentationTransaction = vi.fn(async () => true)
    const result = await executePreparedGeometryFamilyTransaction(
      {
        preparePresentationTarget: vi.fn(async () => ({
          status: 'prepared' as const,
          expectedDeckRevision: fingerprint,
          target: {
            slideId: 'slide-1',
            elementId: 'other',
            expectedType: 'shape' as const,
            expectedFingerprint: fingerprint,
          },
        })),
        executePresentationTransaction,
        cancelPresentationTransaction,
      },
      {
        transactionId: 'tx-scope',
        slideIndex: 0,
        operations: [{ sourceId: 'legacy-other', geometry: { x: 1, y: 1, width: 10, height: 10 } }],
      },
      undefined,
      undefined,
      scope,
    )

    expect(result.receipt).toMatchObject({ status: 'conflict', code: 'target_stale' })
    expect(executePresentationTransaction).not.toHaveBeenCalled()
    expect(cancelPresentationTransaction).toHaveBeenCalledWith('tx-scope')
  })
})
