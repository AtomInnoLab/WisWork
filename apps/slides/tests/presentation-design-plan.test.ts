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

  it('persists legacy planning as a structured draft that requires explicit replan', async () => {
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
    expect(result.output).toContain('Status: draft')
    expect(result.output).toContain('## Deck Acceptance')
    expect(result.output).toContain('WISWORK_PRESENTATION_DESIGN_CONTRACT:')
    expect(saveSidecar).toHaveBeenCalledWith(
      expect.objectContaining({
        topic: 'One coherent story',
        styleSkill: 'Background: #0A0A0A\nTitle: 32pt',
        designMd: expect.stringContaining('Status: draft'),
      }),
    )
  })

  it('blocks mutation when a fresh session opens a non-structured legacy DESIGN.md', async () => {
    const skill = createSlidesSkill({
      getSlides: () =>
        [
          { widthPx: 1280, heightPx: 720, nodes: [] },
          { widthPx: 1280, heightPx: 720, nodes: [] },
        ] as never,
      getCurrent: () => 0,
      getSelectedIds: () => [],
      getPresentationDesignDocument: () => '# DESIGN.md\n\nAccent: #10B981',
      applySlide: () => undefined,
      applyDeck: () => undefined,
      fitWidthPx: 1280,
    })
    skill.buildContext?.()

    const mutation = await skill.executeTool({
      id: 'legacy-mutation',
      name: 'set_slide_background',
      input: { slideIndex: 0, color: '#000000' },
    })
    expect(mutation).toMatchObject({ isError: true, mutated: false })
    expect(mutation.output).toContain('Normalize it into a ready structured contract')
  })

  it('keeps an image-bearing legacy plan as a draft until assets are resolved', async () => {
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
    const pages = [
      {
        title: 'Evidence',
        brief: 'Show the real place',
        layout: 'cover',
        purpose: 'Open',
        visual: 'Documentary photograph',
        acceptance: ['Image supports the claim'],
        density: 'low',
        image_queries: ['real wetland restoration'],
      },
      {
        title: 'Action',
        brief: 'Approve restoration',
        layout: 'statement',
        purpose: 'Close',
        visual: 'Native typography',
        acceptance: ['Action is explicit'],
        density: 'low',
      },
    ]

    const plan = await skill.executeTool({
      id: 'legacy-image-plan',
      name: 'plan_deck',
      input: { core_hook: 'Restore wetlands', style: 'Editorial', pages, prototype_pages: [0, 1] },
    })
    const build = await skill.executeTool({
      id: 'premature-build',
      name: 'build_deck',
      input: {
        pages: pages.map(({ title, brief: body, layout }) => ({ title, body: [body], layout })),
        phase: 'prototype',
        page_indexes: [0, 1],
      },
    })

    expect(plan.output).toContain('Status: draft')
    expect(plan.output).not.toContain('· ready')
    expect(saveSidecar).toHaveBeenCalledWith(
      expect.objectContaining({ designMd: expect.stringContaining('Status: draft') }),
    )
    expect(build).toMatchObject({ isError: true, mutated: false })
    expect(build.output).toContain('draft')
  })

  it('requires session image-search provenance for ready remote assets', async () => {
    const imageUrl = 'https://images.example/wetland.jpg'
    ;(globalThis as unknown as { window: Record<string, unknown> }).window = {
      slidesApi: {
        imageSearch: vi.fn(async () => ({
          images: [{ imageUrl, title: 'Wetland' }],
          method: 'test',
        })),
      },
    }
    const skill = createSlidesSkill({
      getSlides: () => [{ widthPx: 1280, heightPx: 720, nodes: [] }] as never,
      getCurrent: () => 0,
      getSelectedIds: () => [],
      applySlide: () => undefined,
      applyDeck: () => undefined,
      fitWidthPx: 1280,
    })
    const contract = {
      ...modernContract,
      slides: [{ ...modernContract.slides[0], assetIds: ['hero'] }],
      assets: [
        {
          id: 'hero',
          slideNumbers: [1],
          type: 'image',
          role: 'substantive',
          intent: 'Real wetland',
          source: imageUrl,
          crop: '16:9',
          placement: 'full bleed',
          status: 'ready' as const,
          localReference: imageUrl,
        },
      ],
    }

    const premature = await skill.executeTool({
      id: 'unverified-remote',
      name: 'plan_deck',
      input: { contract },
    })
    expect(premature).toMatchObject({ isError: true, mutated: false })
    expect(premature.output).toContain('image_search')

    await skill.executeTool({ id: 'search', name: 'image_search', input: { query: 'wetland' } })
    const accepted = await skill.executeTool({
      id: 'verified-remote',
      name: 'plan_deck',
      input: { contract },
    })
    expect(accepted.isError).not.toBe(true)
    expect(accepted.output).toContain('· ready')
  })

  it('rejects a searched image that is not the ready asset bound to that slide', async () => {
    const assetA = 'https://images.example/contract-a.jpg'
    const searchedB = 'https://images.example/searched-b.jpg'
    ;(globalThis as unknown as { window: Record<string, unknown> }).window = {
      slidesApi: {
        imageSearch: vi.fn(async () => ({
          images: [
            { imageUrl: assetA, title: 'Contract asset' },
            { imageUrl: searchedB, title: 'Different result' },
          ],
          method: 'test',
        })),
      },
    }
    const skill = createSlidesSkill({
      getSlides: () => [{ widthPx: 1280, heightPx: 720, nodes: [] }] as never,
      getCurrent: () => 0,
      getSelectedIds: () => [],
      applySlide: () => undefined,
      applyDeck: () => undefined,
      fitWidthPx: 1280,
    })
    const contract = {
      ...modernContract,
      brief: { ...modernContract.brief, pageCount: 2 },
      prototypePages: [1, 2],
      slides: [
        { ...modernContract.slides[0], assetIds: ['hero'] },
        {
          ...modernContract.slides[0],
          number: 2,
          title: 'Act now',
          layoutFamily: 'statement',
          assetIds: [],
          acceptance: [{ id: 'A2.1', criterion: 'Action is explicit' }],
        },
      ],
      assets: [
        {
          id: 'hero',
          slideNumbers: [1],
          type: 'image',
          role: 'substantive',
          intent: 'Contract hero',
          source: assetA,
          crop: '16:9',
          placement: 'right panel',
          status: 'ready' as const,
          localReference: assetA,
        },
      ],
    }

    await skill.executeTool({ id: 'search-both', name: 'image_search', input: { query: 'hero' } })
    await skill.executeTool({ id: 'plan-bound', name: 'plan_deck', input: { contract } })
    const result = await skill.executeTool({
      id: 'substitute-b',
      name: 'build_deck',
      input: {
        pages: [
          {
            title: 'Nature is infrastructure',
            body: ['One claim'],
            evidence: ['Source'],
            layout: 'cover',
            imageUrl: searchedB,
            imageAlt: 'Wrong searched image',
          },
          {
            title: 'Act now',
            body: ['One claim'],
            evidence: ['Source'],
            layout: 'statement',
          },
        ],
        phase: 'prototype',
        page_indexes: [0, 1],
      },
    })

    expect(result).toMatchObject({ isError: true, mutated: false })
    expect(result.output).toContain('assetIds')
  })

  it('does not treat a native fallback asset source page as a required imageUrl', async () => {
    let slides = [{ widthPx: 1280, heightPx: 720, nodes: [] }] as never[]
    ;(globalThis as unknown as { window: Record<string, unknown> }).window = {
      slidesApi: {
        addSlide: vi.fn(async () => {
          slides = [...slides, { widthPx: 1280, heightPx: 720, nodes: [] } as never]
          return { slides, index: slides.length - 1 }
        }),
      },
    }
    const skill = createSlidesSkill({
      getSlides: () => slides as never,
      getCurrent: () => 0,
      getSelectedIds: () => [],
      applySlide: () => undefined,
      applyDeck: (next) => {
        slides = next as never[]
      },
      executePresentationOperation: vi.fn(async (request) => ({
        receipt: {
          status: 'applied' as const,
          transactionId: request.transactionId,
          resultingDeckRevision: `sha256:${'a'.repeat(64)}`,
          operationCount: request.operations.length,
        },
        authoritativeState: 'fresh' as const,
      })),
      fitWidthPx: 1280,
    })
    const contract = {
      ...modernContract,
      brief: { ...modernContract.brief, pageCount: 2 },
      prototypePages: [1, 2],
      slides: [
        { ...modernContract.slides[0], assetIds: ['native-chart'] },
        {
          ...modernContract.slides[0],
          number: 2,
          title: 'Act now',
          layoutFamily: 'statement',
          acceptance: [{ id: 'A2.1', criterion: 'Action is explicit' }],
        },
      ],
      assets: [
        {
          id: 'native-chart',
          slideNumbers: [1],
          type: 'chart',
          role: 'evidence',
          intent: 'Show the trend as an editable chart',
          source: 'https://example.com/report',
          crop: 'none',
          placement: 'right panel',
          status: 'fallback_ready' as const,
          fallback: 'Native editable chart using cited values',
        },
      ],
    }

    const plan = await skill.executeTool({
      id: 'native-plan',
      name: 'plan_deck',
      input: { contract },
    })
    expect(plan.isError).not.toBe(true)
    const build = await skill.executeTool({
      id: 'native-build',
      name: 'build_deck',
      input: {
        pages: [
          {
            title: 'Nature is infrastructure',
            body: ['One claim'],
            evidence: ['Source'],
            layout: 'cover',
          },
          {
            title: 'Act now',
            body: ['One claim'],
            evidence: ['Source'],
            layout: 'statement',
          },
        ],
        phase: 'prototype',
        page_indexes: [0, 1],
      },
    })
    expect(build.isError).not.toBe(true)
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

  it('persists and exposes a structured draft before asset research is complete', async () => {
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
      id: 'draft-contract',
      name: 'plan_deck',
      input: {
        contract: {
          ...modernContract,
          status: 'draft',
          discovery: {
            questionnaire: ['Audience: independent travellers'],
            openQuestions: ['Confirm image licences'],
            researchNotes: ['Volcanic route candidates collected'],
          },
        },
      },
    })

    expect(result).toMatchObject({ mutated: false })
    expect(result.isError).not.toBe(true)
    expect(result.output).toContain('DESIGN.md · Revision 3 · draft')
    expect(result.output).toContain('NEXT REQUIRED ACTION: run image_search')
    expect(result.output).toContain('Audience: independent travellers')
    expect(result.output).toContain('Volcanic route candidates collected')
    expect(saveSidecar).toHaveBeenCalledWith(
      expect.objectContaining({ designMd: expect.stringContaining('Status: draft') }),
    )
    expect(skill.buildContext?.()).toContain('Status: draft')
  })

  it('records successful image research and candidate assets in the active DESIGN.md draft', async () => {
    ;(globalThis as unknown as { window: Record<string, unknown> }).window = {
      slidesApi: {
        imageSearch: vi.fn(async () => ({
          images: [
            {
              imageUrl: 'https://images.example/borobudur.jpg',
              title: 'Borobudur sunrise',
            },
          ],
          method: 'test',
        })),
      },
    }
    const saveSidecar = vi.fn(async () => undefined)
    const setPresentationDesignContext = vi.fn()
    const skill = createSlidesSkill({
      getSlides: () => [{ widthPx: 1280, heightPx: 720, nodes: [] }] as never,
      getCurrent: () => 0,
      getSelectedIds: () => [],
      applySlide: () => undefined,
      applyDeck: () => undefined,
      saveSidecar,
      setPresentationDesignContext,
      fitWidthPx: 1280,
    })
    await skill.executeTool({
      id: 'draft-before-search',
      name: 'plan_deck',
      input: { contract: { ...modernContract, status: 'draft' } },
    })

    const result = await skill.executeTool({
      id: 'search-and-record',
      name: 'image_search',
      input: { query: 'Borobudur sunrise' },
    })

    expect(result.summary).toContain('DESIGN.md updated')
    expect(saveSidecar).toHaveBeenLastCalledWith(
      expect.objectContaining({
        designMd: expect.stringContaining('https://images.example/borobudur.jpg'),
      }),
    )
    expect(setPresentationDesignContext).toHaveBeenLastCalledWith(
      expect.objectContaining({ designMd: expect.stringContaining('Image search') }),
    )
    expect(skill.buildContext?.()).toContain('Borobudur sunrise')
  })

  it('accepts a lightweight initial draft before expanding the page and asset plans', async () => {
    const skill = createSlidesSkill({
      getSlides: () => [{ widthPx: 1280, heightPx: 720, nodes: [] }] as never,
      getCurrent: () => 0,
      getSelectedIds: () => [],
      applySlide: () => undefined,
      applyDeck: () => undefined,
      fitWidthPx: 1280,
    })

    const result = await skill.executeTool({
      id: 'initial-draft',
      name: 'plan_deck',
      input: {
        contract: {
          ...modernContract,
          status: 'draft',
          brief: { ...modernContract.brief, pageCount: 10 },
          prototypePages: [],
          slides: [],
          assets: [],
          deckAcceptance: [],
        },
      },
    })

    expect(result.isError).not.toBe(true)
    expect(result.output).toContain('DESIGN.md · Revision 3 · draft')
    expect(result.output).toContain('Page Count: 10')
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
        { ...modernContract.slides[0], title: 'Nature — infrastructure' },
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
    const equivalent = await skill.executeTool({
      id: 'build',
      name: 'build_deck',
      input: {
        pages: [
          {
            title: '  Nature\u00a0- infrastructure  ',
            body: ['One claim'],
            evidence: ['Source'],
            layout: ' cover ',
          },
          {
            title: 'Act now',
            body: ['One claim'],
            evidence: ['Source'],
            layout: 'statement',
          },
        ],
        phase: 'prototype',
        page_indexes: [0, 1],
      },
    })
    expect(equivalent.isError).not.toBe(true)
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

  it('rejects semantic title drift from the authoritative contract', async () => {
    const skill = createSlidesSkill({
      getSlides: () => [{ widthPx: 1280, heightPx: 720, nodes: [] }] as never,
      getCurrent: () => 0,
      getSelectedIds: () => [],
      applySlide: () => undefined,
      applyDeck: () => undefined,
      fitWidthPx: 1280,
    })

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
          layoutFamily: 'statement',
          acceptance: [{ id: 'A2.1', criterion: 'Action is explicit' }],
        },
      ],
    }
    await skill.executeTool({ id: 'plan', name: 'plan_deck', input: { contract } })
    const result = await skill.executeTool({
      id: 'drift',
      name: 'build_deck',
      input: {
        pages: [
          { title: 'Nature is optional', body: ['One claim'], layout: 'cover' },
          { title: 'Act now', body: ['Action'], layout: 'statement' },
        ],
        phase: 'prototype',
        page_indexes: [0, 1],
      },
    })

    expect(result).toMatchObject({ isError: true, mutated: false })
    expect(result.output).toContain('must match the active page plan')

    const layoutDrift = await skill.executeTool({
      id: 'layout-drift',
      name: 'build_deck',
      input: {
        pages: [
          { title: 'Nature is infrastructure', body: ['One claim'], layout: 'cards' },
          { title: 'Act now', body: ['Action'], layout: 'statement' },
        ],
        phase: 'prototype',
        page_indexes: [0, 1],
      },
    })
    expect(layoutDrift).toMatchObject({ isError: true, mutated: false })
    expect(layoutDrift.output).toContain('must match the active page plan')
  })

  it('rejects body or evidence drift from the authoritative contract', async () => {
    const skill = createSlidesSkill({
      getSlides: () => [{ widthPx: 1280, heightPx: 720, nodes: [] }] as never,
      getCurrent: () => 0,
      getSelectedIds: () => [],
      applySlide: () => undefined,
      applyDeck: () => undefined,
      fitWidthPx: 1280,
    })
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
          content: ['Approve restoration'],
          evidence: ['Board mandate'],
          layoutFamily: 'statement',
          acceptance: [{ id: 'A2.1', criterion: 'Action is explicit' }],
        },
      ],
    }
    await skill.executeTool({ id: 'plan', name: 'plan_deck', input: { contract } })
    const result = await skill.executeTool({
      id: 'body-drift',
      name: 'build_deck',
      input: {
        pages: [
          {
            title: 'Nature is infrastructure',
            body: ['Contradict the claim'],
            evidence: ['Source'],
            layout: 'cover',
          },
          {
            title: 'Act now',
            body: ['Approve restoration'],
            evidence: ['Board mandate'],
            layout: 'statement',
          },
        ],
        phase: 'prototype',
        page_indexes: [0, 1],
      },
    })
    expect(result).toMatchObject({ isError: true, mutated: false })
    expect(result.output).toContain('body/evidence')
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
