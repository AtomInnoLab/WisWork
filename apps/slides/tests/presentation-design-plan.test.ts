import { describe, expect, it, vi } from 'vitest'
import { renderPresentationDesignContract } from '@wiswork/agent-core'
import { createSlidesSkill } from '../src/renderer/ai/slides-skill'

describe('presentation design plan', () => {
  const modernContract = {
    schemaVersion: 1 as const,
    revision: 3,
    status: 'ready' as const,
    prototypePages: [1],
    brief: {
      topic: 'Nature',
      audience: 'Leaders',
      occasion: 'Review',
      desiredOutcome: 'Approve',
      language: 'English',
      pageCount: 1,
      aspectRatio: '16:9',
      sourceConstraints: [],
    },
    narrative: {
      coreHook: 'Nature is infrastructure',
      opening: 'Risk',
      development: 'Evidence',
      tension: 'Loss',
      resolution: 'Invest',
      closingAction: 'Approve',
    },
    visualSystem: {
      style: 'Editorial nature',
      colors: { accent: '#10B981' },
      typography: { title: '32pt' },
      safeMargin: '64px',
      grid: '12 columns',
      imageTreatment: 'Documentary',
      chartTreatment: 'Direct labels',
      antiPatterns: ['No decorative images'],
    },
    slides: [
      {
        number: 1,
        title: 'Nature is infrastructure',
        role: 'Open',
        claim: 'Invest now',
        content: ['One claim'],
        evidence: ['Source'],
        visualRoute: 'Hero image',
        layoutFamily: 'cover',
        focalVisual: 'Forest',
        density: 'low' as const,
        assetIds: [],
        acceptance: [{ id: 'A1.1', criterion: 'Title is dominant' }],
      },
    ],
    assets: [],
    deckAcceptance: [{ id: 'D1', criterion: 'One claim per slide' }],
  }

  it('places an edited DESIGN.md contract in the next agent-turn context', () => {
    const skill = createSlidesSkill({
      getSlides: () => [{ widthPx: 1280, heightPx: 720, nodes: [] }] as never,
      getCurrent: () => 0,
      getSelectedIds: () => [],
      getPresentationDesignDocument: () => '# DESIGN.md\n\nAccent: #10B981',
      applySlide: () => undefined,
      applyDeck: () => undefined,
      fitWidthPx: 1280,
    })

    expect(skill.buildContext?.()).toContain('Accent: #10B981')
  })

  it('blocks legacy low-level writes on a blank deck until the design plan exists', async () => {
    const skill = createSlidesSkill({
      getSlides: () => [{ widthPx: 1280, heightPx: 720, nodes: [] }] as never,
      getCurrent: () => 0,
      getSelectedIds: () => [],
      applySlide: () => undefined,
      applyDeck: () => undefined,
      fitWidthPx: 1280,
    })
    skill.buildContext?.()

    const result = await skill.executeTool({
      id: 'legacy-write',
      name: 'add_text_box',
      input: { slideIndex: 0, text: 'Bypass the design workflow' },
    })

    expect(result).toMatchObject({ isError: true, mutated: false })
    expect(result.output).toContain('plan_deck')
  })

  it('persists and returns the editable design contract before production', async () => {
    const saveSidecar = vi.fn(async () => undefined)
    const skill = createSlidesSkill({
      getSlides: () => [{ widthPx: 1280, heightPx: 720, nodes: [] }] as never,
      getCurrent: () => 0,
      getSelectedIds: () => [],
      applySlide: () => undefined,
      applyDeck: () => undefined,
      saveSidecar,
      fitWidthPx: 1280,
    })

    const result = await skill.executeTool({
      id: 'plan',
      name: 'plan_deck',
      input: {
        core_hook: 'One coherent story',
        style: 'Background: #0A0A0A\nTitle: 32pt',
        pages: [
          {
            title: 'A conclusion-led cover',
            brief: 'Open with the central claim',
            layout: 'cover',
            purpose: 'Open the story',
            visual: 'One typographic hero',
            acceptance: ['Title is dominant'],
            density: 'low',
          },
        ],
        prototype_pages: [0],
      },
    })

    expect(result.output).toContain('# DESIGN.md')
    expect(saveSidecar).toHaveBeenCalledWith(
      expect.objectContaining({
        topic: 'One coherent story',
        styleSkill: 'Background: #0A0A0A\nTitle: 32pt',
      }),
    )
  })

  it('accepts a ready structured contract, persists full DESIGN.md, and exposes revision context', async () => {
    const saveSidecar = vi.fn(async () => undefined)
    const skill = createSlidesSkill({
      getSlides: () => [{ widthPx: 1280, heightPx: 720, nodes: [] }] as never,
      getCurrent: () => 0,
      getSelectedIds: () => [],
      applySlide: () => undefined,
      applyDeck: () => undefined,
      saveSidecar,
      fitWidthPx: 1280,
    })

    const result = await skill.executeTool({
      id: 'contract',
      name: 'plan_deck',
      input: { contract: modernContract },
    })

    expect(result).toMatchObject({ mutated: false })
    expect(result.isError).not.toBe(true)
    expect(result.output).toContain('DESIGN.md · Revision 3 · ready')
    expect(result.output).toContain('A1.1')
    expect(saveSidecar).toHaveBeenCalledWith(
      expect.objectContaining({
        designMd: expect.stringContaining('Revision: 3'),
      }),
    )
    expect(skill.buildContext?.()).toContain('Revision: 3')
  })

  it('rejects a structured contract that is not ready for production', async () => {
    const skill = createSlidesSkill({
      getSlides: () => [{ widthPx: 1280, heightPx: 720, nodes: [] }] as never,
      getCurrent: () => 0,
      getSelectedIds: () => [],
      applySlide: () => undefined,
      applyDeck: () => undefined,
      fitWidthPx: 1280,
    })

    const result = await skill.executeTool({
      id: 'draft-contract',
      name: 'plan_deck',
      input: { contract: { ...modernContract, status: 'draft' } },
    })

    expect(result).toMatchObject({ isError: true, mutated: false })
    expect(result.output).toContain('status must be ready')
  })

  it('moves a contract to producing on build and verifies only after final verify_slides', async () => {
    let slides = [{ widthPx: 1280, heightPx: 720, nodes: [] }] as never[]
    ;(globalThis as unknown as { window: Record<string, unknown> }).window = {
      slidesApi: {
        addSlide: vi.fn(async () => {
          slides = [...slides, { widthPx: 1280, heightPx: 720, nodes: [] } as never]
          return { slides, index: slides.length - 1 }
        }),
      },
    }
    const contract = {
      ...modernContract,
      brief: { ...modernContract.brief, pageCount: 2 },
      prototypePages: [1, 2],
      slides: [
        modernContract.slides[0],
        {
          ...modernContract.slides[0],
          number: 2,
          title: 'Act now',
          role: 'Close',
          claim: 'Approve the plan',
          layoutFamily: 'statement',
          acceptance: [{ id: 'A2.1', criterion: 'Action is explicit' }],
        },
      ],
    }
    const saveSidecar = vi.fn(async () => undefined)
    const skill = createSlidesSkill({
      getSlides: () => slides as never,
      getCurrent: () => 0,
      getSelectedIds: () => [],
      applySlide: () => undefined,
      applyDeck: (next) => {
        slides = next as never[]
      },
      captureSlideScreenshot: vi.fn(async () => ({ base64: 'AA==', mime: 'image/png' })),
      reviewPresentationScreenshot: vi.fn(async () => true),
      executePresentationOperation: vi.fn(async (request) => ({
        receipt: {
          status: 'applied' as const,
          transactionId: request.transactionId,
          resultingDeckRevision: `sha256:${'a'.repeat(64)}`,
          operationCount: request.operations.length,
        },
        authoritativeState: 'fresh' as const,
      })),
      saveSidecar,
      fitWidthPx: 1280,
    })

    await skill.executeTool({ id: 'plan', name: 'plan_deck', input: { contract } })
    await skill.executeTool({
      id: 'build',
      name: 'build_deck',
      input: {
        pages: [
          { title: 'Nature is infrastructure', body: ['One claim'], layout: 'cover' },
          { title: 'Act now', body: ['Approve the plan'], layout: 'statement' },
        ],
        phase: 'prototype',
        page_indexes: [0, 1],
      },
    })
    expect(saveSidecar).toHaveBeenCalledWith(
      expect.objectContaining({
        designMd: expect.stringContaining('Status: producing'),
      }),
    )
    const early = await skill.executeTool({ id: 'early', name: 'verify_slides', input: {} })
    expect(early).toMatchObject({ isError: true })

    await skill.executeTool({ id: 'shot-1', name: 'screenshot_slide', input: { slideIndex: 0 } })
    await skill.executeTool({ id: 'shot-2', name: 'screenshot_slide', input: { slideIndex: 1 } })
    const verified = await skill.executeTool({ id: 'verify', name: 'verify_slides', input: {} })

    expect(verified.output).toContain('Revision 3 · verified')
    expect(saveSidecar).toHaveBeenCalledWith(
      expect.objectContaining({
        designMd: expect.stringContaining('Status: verified'),
      }),
    )
  })

  it('restores a producing contract from DESIGN.md and requires fresh screenshot review', async () => {
    const producing = { ...modernContract, status: 'producing' as const }
    const skill = createSlidesSkill({
      getSlides: () => [{ widthPx: 1280, heightPx: 720, nodes: [] }] as never,
      getCurrent: () => 0,
      getSelectedIds: () => [],
      getPresentationDesignDocument: () => renderPresentationDesignContract(producing),
      applySlide: () => undefined,
      applyDeck: () => undefined,
      fitWidthPx: 1280,
    })

    expect(skill.buildContext?.()).toContain('Status: producing')
    await expect(
      skill.executeTool({ id: 'verify', name: 'verify_slides', input: {} }),
    ).resolves.toMatchObject({
      isError: true,
      output: expect.stringContaining('Only 0 of 1 planned slides have been built'),
    })
  })

  it('requires replanning after visible contract edits invalidate the structured snapshot', async () => {
    let designMd = renderPresentationDesignContract(modernContract)
    const skill = createSlidesSkill({
      getSlides: () => [{ widthPx: 1280, heightPx: 720, nodes: [] }] as never,
      getCurrent: () => 0,
      getSelectedIds: () => [],
      getPresentationDesignDocument: () => designMd,
      applySlide: () => undefined,
      applyDeck: () => undefined,
      fitWidthPx: 1280,
    })
    skill.buildContext?.()
    designMd = '# DESIGN.md\n\nStatus: draft\nRevision: 4\n\nUser-edited contract'
    skill.buildContext?.()

    await expect(
      skill.executeTool({
        id: 'blocked-edit',
        name: 'set_element_text',
        input: { slideIndex: 0, sourceId: 'title', paragraphs: [{ runs: [{ text: 'Changed' }] }] },
      }),
    ).resolves.toMatchObject({
      isError: true,
      output: expect.stringContaining('Normalize it into a ready structured contract'),
    })
  })
})
