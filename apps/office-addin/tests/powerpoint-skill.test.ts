import { afterEach, describe, expect, it, vi } from 'vitest'
import JSZip from 'jszip'
import {
  extractPresentationDesignContract,
  PRESENTATION_DESIGN_CONTRACT_SCHEMA,
} from '@wiswork/agent-core'
import { createStructuredProposalController } from '../src/agent/proposal-controller.js'
import {
  BrowserPowerPointAdapter,
  type PowerPointAdapter,
} from '../src/skills/powerpoint/browser-powerpoint-adapter.js'
import { createPowerPointSkill } from '../src/skills/powerpoint/powerpoint-skill.js'
import { editPowerPointPackage } from '../src/skills/powerpoint/powerpoint-package.js'

const png =
  'iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAQAAAC1HAwCAAAAC0lEQVR42mP8/x8AAwMCAO+/DPsAAAAASUVORK5CYII='

function adapter(overrides: Partial<PowerPointAdapter> = {}): PowerPointAdapter {
  return {
    getPresentationState: vi.fn().mockResolvedValue({
      slideCount: 1,
      selectedSlideIndexes: [0],
      api: { v12: true, v14: true, v15: true, v18: true, v110: true },
    }),
    inspectSlideMasters: vi.fn().mockResolvedValue({
      masters: [
        {
          id: 'master-1',
          name: 'Main',
          background: { type: 'Solid', color: '#FFFFFF', transparency: 0 },
          themeColors: { Light1: '#FFFFFF', Dark1: '#000000' },
          layouts: [
            {
              id: 'layout-1',
              name: 'Title',
              isMasterBackgroundFollowed: true,
              areBackgroundGraphicsHidden: false,
              background: { type: 'Solid' },
            },
          ],
        },
      ],
    }),
    executeMasterOperations: vi.fn().mockResolvedValue(undefined),
    screenshotSlide: vi.fn().mockResolvedValue({ mime: 'image/png', base64: png }),
    listSlideShapes: vi.fn().mockResolvedValue({
      slideId: 'slide-1',
      slideIndex: 0,
      shapes: [
        { id: '2', name: 'Title', type: 'TextBox', left: 10, top: 20, width: 200, height: 40 },
      ],
    }),
    readSlideText: vi.fn().mockResolvedValue({
      slideId: 'slide-1',
      shapeId: '2',
      text: 'Hello',
      paragraphs: ['Hello'],
    }),
    verifySlides: vi.fn().mockResolvedValue({
      slideWidth: 960,
      slideHeight: 540,
      slides: [],
    }),
    snapshotSlide: vi.fn().mockResolvedValue({ slideId: 'slide-1', fingerprint: 'slide-1:1' }),
    editSlideText: vi.fn().mockResolvedValue(undefined),
    duplicateSlide: vi.fn().mockResolvedValue({ slideId: 'slide-copy' }),
    exportSlidePackage: vi.fn().mockRejectedValue(new Error('office_api_unsupported')),
    replaceSlidePackage: vi.fn().mockRejectedValue(new Error('office_api_unsupported')),
    executeDeclarative: vi.fn().mockRejectedValue(new Error('office_api_unsupported')),
    ...overrides,
  }
}

const call = (name: string, input: Record<string, unknown> = {}) => ({ id: 'call-1', name, input })

const modernContract = (overrides: Record<string, unknown> = {}) => ({
  schemaVersion: 1,
  revision: 2,
  status: 'ready',
  prototypePages: [1],
  brief: {
    topic: 'AI onboarding',
    audience: 'New hires',
    occasion: 'Orientation',
    desiredOutcome: 'Adopt the workflow',
    language: 'English',
    pageCount: 1,
    aspectRatio: '16:9',
    sourceConstraints: [],
  },
  narrative: {
    coreHook: 'Reach useful output in seven days',
    opening: 'Start with the goal',
    development: 'Show the workflow',
    tension: 'Avoid common failure modes',
    resolution: 'Use a guided first week',
    closingAction: 'Start today',
  },
  visualSystem: {
    style: 'Editorial blue system',
    colors: { accent: '#2367E8' },
    typography: { title: 'Aptos Display 32pt' },
    safeMargin: '48pt',
    grid: '12 columns',
    imageTreatment: 'Natural documentary photography',
    chartTreatment: 'Direct labels',
    antiPatterns: ['repetitive cards'],
  },
  slides: [
    {
      number: 1,
      title: 'First useful result',
      role: 'Open',
      claim: 'New hires can contribute in seven days',
      content: ['A guided sequence'],
      evidence: [],
      visualRoute: 'Hero statement',
      layoutFamily: 'cover',
      focalVisual: 'One strong headline',
      density: 'low',
      assetIds: [],
      acceptance: [{ id: 'A1.1', criterion: 'Headline is dominant' }],
    },
  ],
  assets: [],
  deckAcceptance: [{ id: 'D1', criterion: 'The story is coherent' }],
  ...overrides,
})

describe('PowerPoint compatibility skill', () => {
  it('returns actual font readback for targeted repair without losing text on unsupported hosts', async () => {
    const port = adapter({
      readShapeTextStyle: vi
        .fn()
        .mockResolvedValue({ fontFamily: 'Aptos', fontSize: 24, color: '#FFFFFF' }),
    })
    const skill = createPowerPointSkill({
      adapter: port,
      proposals: createStructuredProposalController(),
    })
    const read = () => skill.executeTool(call('read_slide_text', { slide_index: 0, shape_id: '2' }))
    expect(JSON.parse((await read()).output)).toMatchObject({
      text: 'Hello',
      textStyle: { fontFamily: 'Aptos', fontSize: 24 },
    })
    vi.mocked(port.readShapeTextStyle!).mockRejectedValue(new Error('office_api_unsupported'))
    expect(JSON.parse((await read()).output)).toMatchObject({ text: 'Hello' })
  })
  it('keeps measured point dimensions in context without inventing a default canvas', async () => {
    const state = { slideCount: 2, selectedSlideIndexes: [], api: {} }
    const port = adapter({
      getPresentationState: vi.fn().mockResolvedValue({
        ...state,
        slideWidth: 960,
        slideHeight: 540,
        coordinateUnit: 'pt',
      }),
    })
    const skill = createPowerPointSkill({
      adapter: port,
      proposals: createStructuredProposalController(),
    })
    expect(skill.buildContext?.()).not.toContain('960')
    await skill.executeTool(call('get_presentation_state'))
    expect(skill.buildContext?.()).toContain('"slideWidth":960')
    expect(skill.buildContext?.()).toContain('"coordinateUnit":"pt"')
    vi.mocked(port.getPresentationState).mockResolvedValue(state as any)
    await skill.executeTool(call('get_presentation_state'))
    expect(skill.buildContext?.()).not.toContain('960')
  })
  it('binds an image proposal to its own page, not the previous text-edit page', async () => {
    const base = modernContract()
    const proposals = createStructuredProposalController()
    const skill = createPowerPointSkill({ adapter: adapter(), proposals })
    await skill.executeTool(
      call('plan_deck', {
        contract: modernContract({
          prototypePages: [1, 2],
          brief: { ...base.brief, pageCount: 2 },
          slides: [1, 2].map((number) => ({
            ...base.slides[0],
            number,
            acceptance: [{ id: `A${number}.1`, criterion: 'Readable' }],
          })),
        }),
      }),
    )
    await skill.executeTool(
      call('edit_slide_text', { slide_index: 0, shape_id: '2', text: 'Hello' }),
    )
    await proposals.confirm(proposals.pending()!.id)
    await skill.executeTool(call('screenshot_slide', { slide_index: 0 }))
    await skill.executeTool(
      call('review_slide_screenshot', { slide_index: 0, acceptance_ids: ['A1.1'], passed: true }),
    )
    const picture = proposals.propose({
      operation: 'insert_web_image',
      toolName: 'insert_web_image',
      title: 'Picture',
      preview: {},
      impact: { host: 'powerpoint', targets: ['slide-2'], count: 1 },
      fingerprint: 's2',
      powerPointMutation: { indexes: [1], scaffold: false },
      validate: () => true,
      execute: () => undefined,
      verify: () => undefined,
    })
    await proposals.confirm(picture.id)
    const progress = skill.buildContext?.().split('<presentation review progress>')[1] ?? ''
    expect(progress).toContain('"slide_index":1')
    expect(progress).not.toContain('"slide_index":0')
    await skill.executeTool(call('screenshot_slide', { slide_index: 1 }))
    expect(
      (
        await skill.executeTool(
          call('review_slide_screenshot', {
            slide_index: 1,
            acceptance_ids: ['A2.1'],
            passed: true,
          }),
        )
      ).isError,
    ).not.toBe(true)
  })
  it('promotes a fully populated draft to ready instead of treating the transition as a replay', async () => {
    const proposals = createStructuredProposalController()
    const skill = createPowerPointSkill({ adapter: adapter(), proposals })
    await skill.executeTool(call('plan_deck', { contract: modernContract({ status: 'draft' }) }))
    const ready = await skill.executeTool(call('plan_deck', { contract: modernContract() }))
    expect(JSON.parse(ready.output).status).toBe('ready')
  })
  it('advertises only caller-owned planning statuses and explains the screenshot review handoff', () => {
    const skill = createPowerPointSkill({
      adapter: adapter(),
      proposals: createStructuredProposalController(),
    })
    const schema = skill.tools.find((tool) => tool.name === 'plan_deck')!.inputSchema as any
    expect(schema.properties.contract.properties.status.enum).toEqual(['draft', 'ready'])
    expect(skill.systemPrompt).toContain('review_slide_screenshot')
    expect(skill.systemPrompt).toContain('Do not resubmit plan_deck to record a review')
  })

  it('exposes uncapped asset inventory, slide references, and legacy image queries', () => {
    const skill = createPowerPointSkill({
      adapter: adapter(),
      proposals: createStructuredProposalController(),
    })
    const schema = skill.tools.find((tool) => tool.name === 'plan_deck')!.inputSchema as any
    const contract = schema.properties.contract.properties
    const shared = PRESENTATION_DESIGN_CONTRACT_SCHEMA.properties as any
    expect(contract.assets).toBe(shared.assets)
    expect(contract.slides).toBe(shared.slides)
    expect(contract.assets).not.toHaveProperty('maxItems')
    expect(contract.slides.items.properties.assetIds).not.toHaveProperty('maxItems')
    expect(schema.properties.pages.items.properties.image_queries).not.toHaveProperty('maxItems')
    expect(schema.properties.pages.items.properties.image_queries.items).toEqual({
      type: 'string',
      minLength: 1,
      maxLength: 200,
    })
  })

  it('preserves produced pages and pending review when the unchanged ready plan is resubmitted', async () => {
    const proposals = createStructuredProposalController()
    const skill = createPowerPointSkill({ adapter: adapter(), proposals })
    const contract = modernContract()
    await skill.executeTool(call('plan_deck', { contract }))
    await skill.executeTool(
      call('edit_slide_text', { slide_index: 0, shape_id: '2', text: 'Hello' }),
    )
    await proposals.confirm(proposals.pending()!.id)
    await skill.executeTool(call('screenshot_slide', { slide_index: 0 }))
    expect(skill.reviewFinalResponse?.({ text: 'Done', mutated: true })).toContain(
      'review_slide_screenshot',
    )
    const repeated = await skill.executeTool(call('plan_deck', { contract }))
    expect(JSON.parse(repeated.output).status).toBe('producing')
    expect(
      await skill.executeTool(
        call('review_slide_screenshot', {
          slide_index: 0,
          acceptance_ids: ['A1.1'],
          passed: true,
        }),
      ),
    ).not.toHaveProperty('isError', true)
    expect(await skill.executeTool(call('verify_slides'))).not.toHaveProperty('isError', true)
    expect(skill.buildContext?.()).toContain('"status":"verified"')
  })

  it('returns actionable review recovery and requires a fresh screenshot without rewriting the contract', async () => {
    const proposals = createStructuredProposalController()
    const skill = createPowerPointSkill({ adapter: adapter(), proposals })
    await skill.executeTool(call('plan_deck', { contract: modernContract() }))
    await skill.executeTool(
      call('edit_slide_text', { slide_index: 0, shape_id: '2', text: 'Hello' }),
    )
    await proposals.confirm(proposals.pending()!.id)
    expect(
      await skill.executeTool(
        call('review_slide_screenshot', {
          slide_index: 0,
          acceptance_ids: ['A1.1'],
          passed: true,
        }),
      ),
    ).toMatchObject({ isError: true, output: 'design_contract_screenshot_required' })
    const result = await skill.executeTool(
      call('plan_deck', { contract: modernContract({ status: 'producing' }) }),
    )
    expect(JSON.parse(result.output)).toMatchObject({
      error: 'design_contract_invalid_status',
      allowedStatuses: ['draft', 'ready'],
      nextTool: 'screenshot_slide',
      pendingReviews: [{ slide_index: 0, acceptance_ids: ['A1.1'], needsScreenshot: true }],
    })
    const shot = await skill.executeTool(call('screenshot_slide', { slide_index: 0 }))
    expect(JSON.parse(shot.output)).toMatchObject({ nextTool: 'review_slide_screenshot' })
    expect(skill.buildContext?.()).toContain('"pendingReviews"')
  })

  it('shares the design prototype and batch verification workflow with desktop Slides', () => {
    const skill = createPowerPointSkill({
      adapter: adapter(),
      proposals: createStructuredProposalController(),
    })
    expect(skill.systemPrompt).toContain('DESIGN.md')
    expect(skill.systemPrompt).toContain('representative content page')
    expect(skill.systemPrompt).toContain('2–3 slides')
  })

  it('exposes the same plan-before-build workflow as desktop Slides', async () => {
    const imageQueries = Array.from({ length: 5 }, (_, index) => `onboarding team scene ${index}`)
    const skill = createPowerPointSkill({
      adapter: adapter(),
      proposals: createStructuredProposalController(),
    })
    expect(skill.tools.map((tool) => tool.name)).toContain('plan_deck')
    expect(skill.systemPrompt).toContain('inspect the presentation before planning')
    expect(skill.systemPrompt).toContain('plan_deck before the first mutation')
    expect(skill.systemPrompt).toContain('must call ask_clarification')
    expect(skill.systemPrompt).toContain('Never replace that tool call with prose questions')
    expect(skill.systemPrompt).toContain('verify_slides after the approved build')
    expect(skill.systemPrompt).toContain('user-visible progress note before every tool batch')

    const imagePlan = await skill.executeTool(
      call('plan_deck', {
        core_hook: 'New hires reach their first useful result in seven days',
        style: 'Clear blue training system with one idea per slide',
        pages: [
          {
            title: 'Welcome',
            type: 'cover',
            brief: 'Set expectations for the first week',
            layout: 'hero_statement',
            purpose: 'Open the story',
            visual: 'One welcoming team photograph',
            acceptance: ['Title is dominant', 'Image supports the message'],
            density: 'low',
            image_queries: imageQueries,
          },
          {
            title: 'Your first seven days',
            type: 'content',
            brief: 'Show the onboarding milestones',
            layout: 'timeline',
            purpose: 'Explain the sequence',
            visual: 'One horizontal milestone timeline',
            acceptance: ['Milestones scan left to right'],
            density: 'medium',
            image_queries: [],
          },
        ],
        prototype_pages: [0, 1],
      }),
    )
    expect(imagePlan.isError).not.toBe(true)
    expect(imagePlan).toMatchObject({
      mutated: false,
      summary: 'Planned 2 slides',
      output: expect.stringContaining('New hires reach their first useful result in seven days'),
    })
    const contract = extractPresentationDesignContract(JSON.parse(imagePlan.output).designMd)
    expect(contract?.assets.map((asset) => asset.intent)).toEqual(imageQueries)
    expect(contract?.slides[0]?.assetIds).toEqual(contract?.assets.map((asset) => asset.id))
    const result = await skill.executeTool(
      call('plan_deck', {
        core_hook: 'One visual system',
        style: 'Background: #0A0A0A',
        pages: [
          {
            title: 'Cover',
            brief: 'Opening',
            layout: 'cover',
            purpose: 'Open',
            visual: 'One title composition',
            acceptance: ['Clear hierarchy'],
            density: 'low',
          },
        ],
        prototype_pages: [0],
      }),
    )
    expect(result.output).toContain('# DESIGN.md')
  })

  it('accepts a ready structured contract, renders its full snapshot, and keeps it in context', async () => {
    const proposals = createStructuredProposalController()
    const skill = createPowerPointSkill({
      adapter: adapter(),
      proposals,
    })
    const schema = skill.tools.find((tool) => tool.name === 'plan_deck')!.inputSchema as any
    expect(schema.properties.contract).toBeDefined()
    expect(schema.anyOf).toEqual(
      expect.arrayContaining([
        { required: ['contract'] },
        { required: ['core_hook', 'style', 'pages', 'prototype_pages'] },
      ]),
    )

    const result = await skill.executeTool(call('plan_deck', { contract: modernContract() }))
    const output = JSON.parse(result.output)
    expect(output).toMatchObject({ status: 'ready', revision: 2 })
    expect(output.designMd).toContain('## Deck Acceptance')
    expect(output.designMd).toContain('A1.1')
    expect(skill.buildContext?.()).toContain('"revision":2')
    expect(skill.buildContext?.()).toContain('"status":"ready"')

    await expect(skill.executeTool(call('verify_slides'))).resolves.toMatchObject({
      isError: true,
      output: expect.stringContaining('"error":"design_contract_production_incomplete"'),
    })
    await skill.executeTool(
      call('edit_slide_text', { slide_index: 0, shape_id: '2', text: 'Hello' }),
    )
    expect(skill.buildContext?.()).toContain('"status":"ready"')
    await proposals.confirm(proposals.pending()!.id)
    expect(skill.buildContext?.()).toContain('"status":"producing"')
    const screenshot = JSON.parse(
      (await skill.executeTool(call('screenshot_slide', { slide_index: 0 }))).output,
    )
    expect(screenshot).toMatchObject({
      designRevision: 2,
      designStatus: 'producing',
      acceptanceIds: ['A1.1'],
    })
    const review = await skill.executeTool(
      call('review_slide_screenshot', {
        slide_index: 0,
        acceptance_ids: ['A1.1'],
        passed: true,
      }),
    )
    expect(review.isError).not.toBe(true)
    const verification = JSON.parse((await skill.executeTool(call('verify_slides'))).output)
    expect(verification).toMatchObject({ status: 'verified', revision: 2 })
    expect(verification.designMd).toContain('Status: verified')
    expect(skill.buildContext?.()).toContain('"status":"verified"')
  })

  it('records an incomplete structured draft before research and preserves legacy planning', async () => {
    const skill = createPowerPointSkill({
      adapter: adapter(),
      proposals: createStructuredProposalController(),
    })
    const invalid = modernContract({
      brief: { ...(modernContract().brief as object), audience: '' },
    })
    await expect(
      skill.executeTool(call('plan_deck', { contract: invalid })),
    ).resolves.toMatchObject({
      isError: true,
      output: expect.stringContaining('brief.audience is required'),
    })
    await expect(
      skill.executeTool(
        call('plan_deck', {
          contract: modernContract({
            status: 'draft',
            discovery: {
              questionnaire: ['Audience: independent travellers'],
              openQuestions: ['Confirm image licences'],
              researchNotes: ['Volcanic route candidates collected'],
            },
          }),
        }),
      ),
    ).resolves.toMatchObject({
      mutated: false,
      output: expect.stringContaining('Volcanic route candidates collected'),
    })
    expect(skill.buildContext?.()).toContain('"status":"draft"')

    await expect(
      skill.executeTool(
        call('plan_deck', {
          contract: modernContract({
            status: 'draft',
            brief: { ...(modernContract().brief as object), pageCount: 10 },
            prototypePages: [],
            slides: [],
            assets: [],
            deckAcceptance: [],
          }),
        }),
      ),
    ).resolves.toMatchObject({
      mutated: false,
      output: expect.stringContaining('"status":"draft"'),
    })

    await expect(
      skill.executeTool(
        call('plan_deck', {
          core_hook: 'Legacy hook',
          style: 'Legacy style',
          pages: [
            {
              title: 'Legacy page',
              brief: 'Legacy brief',
              layout: 'cover',
              purpose: 'Open',
              visual: 'Headline',
              acceptance: ['Readable'],
              density: 'low',
            },
          ],
          prototype_pages: [0],
        }),
      ),
    ).resolves.toMatchObject({ mutated: false, output: expect.stringContaining('# DESIGN.md') })
  })

  it('explains every unresolved asset before retrying ready and preserves the active draft', async () => {
    const skill = createPowerPointSkill({
      adapter: adapter(),
      proposals: createStructuredProposalController(),
    })
    const base = modernContract()
    const assets = Array.from({ length: 4 }, (_, index) => ({
      id: `image-${index}`,
      slideNumbers: [index + 1],
      type: 'image',
      role: 'evidence',
      intent: 'Show the evidence',
      source: `https://sources.example/page-${index}`,
      crop: '16:9',
      placement: 'right',
      status: 'validated',
      ...(index === 3 ? {} : { localReference: `https://images.example/${index}.jpg` }),
    }))
    const draft = modernContract({
      status: 'draft',
      prototypePages: [1, 2, 3],
      brief: { ...base.brief, pageCount: 6 },
      slides: Array.from({ length: 6 }, (_, index) => ({
        ...base.slides[0],
        number: index + 1,
        assetIds: index === 1 ? [] : [`image-${index % 4}`],
        acceptance: [{ id: `A${index + 1}.1`, criterion: 'Readable' }],
      })),
      assets,
    })
    await skill.executeTool(call('plan_deck', { contract: draft }))
    const previousContext = skill.buildContext?.()
    const rejected = await skill.executeTool(
      call('plan_deck', { contract: { ...draft, status: 'ready' } }),
    )
    expect(rejected.isError).toBe(true)
    for (const index of [0, 2, 3, 4, 5])
      expect(rejected.output).toContain(`slides[${index}].assetIds must reference a ready asset`)
    for (let index = 0; index < 4; index++) {
      expect(rejected.output).toContain(`assets[${index}] must be ready or fallback_ready`)
      expect(rejected.output).toContain(`"id":"image-${index}","status":"validated"`)
    }
    expect(rejected.output).toContain('"localReference":false')
    expect(rejected.output).toContain('"missingForReady":["localReference"]')
    expect(rejected.output).toContain('Resubmit the full corrected contract with plan_deck')
    expect(skill.buildContext?.()).toBe(previousContext)
    expect(
      await skill.executeTool(call('set_slide_background', { slide_index: 0, color: '#FFFFFF' })),
    ).toMatchObject({ isError: true })

    const repaired = {
      ...draft,
      status: 'ready',
      assets: assets.map((asset, index) => ({
        ...asset,
        status: 'ready',
        localReference: `https://images.example/${index}.jpg`,
      })),
    }
    expect(await skill.executeTool(call('plan_deck', { contract: repaired }))).toMatchObject({
      mutated: false,
      output: expect.stringContaining('"status":"ready"'),
    })
    expect(skill.buildContext?.()).toContain('"localReference":"https://images.example/3.jpg"')
  })

  it('allows background repair of a reviewed page even while three other pages await review', async () => {
    const base = modernContract()
    const contract = modernContract({
      prototypePages: [1, 2, 3],
      brief: { ...base.brief, pageCount: 4 },
      slides: [1, 2, 3, 4].map((number) => ({
        ...base.slides[0],
        number,
        acceptance: [{ id: `A${number}.1`, criterion: 'Readable' }],
      })),
    })
    const proposals = createStructuredProposalController()
    let color = '#FFFFFF'
    const fake = adapter({
      getPresentationState: vi
        .fn()
        .mockResolvedValue({ slideCount: 4, selectedSlideIndexes: [0], api: {} }),
      readSlideBackground: vi.fn(async () => ({
        slideId: 'slide-1',
        type: 'Solid',
        backgroundColor: color,
        transparency: 0,
      })),
      setSlideBackground: vi.fn(async (_index, next) => {
        color = next
      }),
    })
    const skill = createPowerPointSkill({ adapter: fake, proposals })
    await skill.executeTool(call('plan_deck', { contract }))
    await skill.executeTool(call('get_presentation_state'))
    for (const slide_index of [0, 1, 2, 3]) {
      expect(
        (
          await skill.executeTool(
            call('edit_slide_text', { slide_index, shape_id: '2', text: 'Hello' }),
          )
        ).isError,
      ).not.toBe(true)
      await proposals.confirm(proposals.pending()!.id)
      if (slide_index === 0) {
        await skill.executeTool(call('screenshot_slide', { slide_index }))
        await skill.executeTool(
          call('review_slide_screenshot', { slide_index, acceptance_ids: ['A1.1'], passed: true }),
        )
      }
    }
    expect(
      (await skill.executeTool(call('set_slide_background', { slide_index: 0, color: '#111111' })))
        .isError,
    ).not.toBe(true)
    await proposals.confirm(proposals.pending()!.id)
    expect(color).toBe('#111111')
    expect(
      await skill.executeTool(
        call('review_slide_screenshot', { slide_index: 0, acceptance_ids: ['A1.1'], passed: true }),
      ),
    ).toMatchObject({ isError: true, output: 'design_contract_screenshot_required' })
  })

  it.each([true, false])(
    'invalidates completed review after a repair (confirmed: %s)',
    async (confirmed) => {
      const proposals = createStructuredProposalController()
      const fake = adapter()
      const skill = createPowerPointSkill({ adapter: fake, proposals })
      await skill.executeTool(call('plan_deck', { contract: modernContract() }))
      await skill.executeTool(
        call('edit_slide_text', { slide_index: 0, shape_id: '2', text: 'Hello' }),
      )
      await proposals.confirm(proposals.pending()!.id)
      await skill.executeTool(call('screenshot_slide', { slide_index: 0 }))
      await skill.executeTool(
        call('review_slide_screenshot', { slide_index: 0, acceptance_ids: ['A1.1'], passed: true }),
      )
      await skill.executeTool(call('verify_slides'))
      expect(skill.buildContext?.()).toContain('"status":"verified"')
      await skill.executeTool(
        call('edit_slide_text', { slide_index: 0, shape_id: '2', text: 'Hello' }),
      )
      if (!confirmed)
        vi.mocked(fake.editSlideText).mockImplementationOnce(async () => {
          vi.mocked(fake.readSlideText).mockResolvedValue({
            slideId: 'slide-1',
            shapeId: '2',
            text: 'Different',
            paragraphs: ['Different'],
          })
        })
      const confirmation = proposals.confirm(proposals.pending()!.id)
      if (confirmed) await confirmation
      else await expect(confirmation).rejects.toThrow('office_verify_failed')
      expect(skill.buildContext?.()).toContain('"status":"producing"')
      expect(skill.buildContext?.()).toContain('"needsScreenshot":true')
      expect(skill.reviewFinalResponse?.({ text: 'Done', mutated: true })).toContain(
        'screenshot_slide',
      )
      expect((await skill.executeTool(call('verify_slides'))).isError).toBe(true)
    },
  )

  it('promotes an applied-unverified prototype after screenshot review and makes repeat review idempotent', async () => {
    const contract = modernContract()
    const fake = adapter({
      editSlideText: vi.fn().mockRejectedValueOnce(new Error('office_applied_unverified')),
    })
    const proposals = createStructuredProposalController()
    const skill = createPowerPointSkill({ adapter: fake, proposals })
    await skill.executeTool(call('plan_deck', { contract }))
    await skill.executeTool(
      call('edit_slide_text', { slide_index: 0, shape_id: '2', text: 'Prototype' }),
    )
    await proposals.confirm(proposals.pending()!.id)
    expect(
      await skill.executeTool(call('screenshot_slide', { slide_index: 0 })),
    ).not.toHaveProperty('isError', true)
    const reviewInput = {
      slide_index: 0,
      acceptance_ids: ['A1.1'],
      passed: true,
    }
    const firstReview = await skill.executeTool(call('review_slide_screenshot', reviewInput))
    expect(firstReview).not.toHaveProperty('isError', true)
    expect(
      await skill.executeTool(
        call('review_slide_screenshot', { ...reviewInput, acceptance_ids: ['wrong'] }),
      ),
    ).toMatchObject({ isError: true, output: 'design_contract_acceptance_mismatch' })
    expect(await skill.executeTool(call('review_slide_screenshot', reviewInput))).toMatchObject({
      output: expect.stringContaining('"status":"already_reviewed"'),
    })
    await expect(skill.executeTool(call('verify_slides'))).resolves.not.toHaveProperty(
      'isError',
      true,
    )
    expect(skill.buildContext?.()).toContain('"status":"verified"')
  })

  it('enforces prototype-first production and rejects host verification defects', async () => {
    const base = modernContract()
    const first = (base.slides as Array<Record<string, unknown>>)[0]!
    const slides = [1, 2, 3, 4].map((number) => ({
      ...first,
      number,
      title: `Slide ${number}`,
      acceptance: [{ id: `A${number}.1`, criterion: 'Readable' }],
    }))
    const contract = modernContract({
      prototypePages: [1, 2, 3],
      brief: { ...(base.brief as object), pageCount: 4 },
      slides,
    })
    const proposals = createStructuredProposalController()
    const fake = adapter({
      getPresentationState: vi.fn().mockResolvedValue({
        slideCount: 4,
        selectedSlideIndexes: [0],
        api: { v12: true },
      }),
      verifySlides: vi.fn().mockResolvedValue({
        slideWidth: 960,
        slideHeight: 540,
        slides: [
          {
            slideId: 'slide-1',
            slideIndex: 0,
            shapes: [],
            shapesTruncated: false,
            overflows: [],
            overlaps: [{ shapeAId: '1', shapeBId: '2', overlapX: 10, overlapY: 10 }],
            overlapsTruncated: false,
          },
        ],
      }),
    })
    const skill = createPowerPointSkill({ adapter: fake, proposals })
    await skill.executeTool(call('plan_deck', { contract }))
    await skill.executeTool(call('get_presentation_state'))

    await expect(
      skill.executeTool(call('edit_slide_text', { slide_index: 3, shape_id: '2', text: 'Hello' })),
    ).resolves.toMatchObject({
      isError: true,
      output: 'design_contract_prototype_required',
    })
    expect(proposals.pending()).toBeUndefined()

    await skill.executeTool(
      call('edit_slide_text', { slide_index: 0, shape_id: '2', text: 'Hello' }),
    )
    await proposals.confirm(proposals.pending()!.id)
    await skill.executeTool(call('screenshot_slide', { slide_index: 0 }))
    await skill.executeTool(
      call('review_slide_screenshot', {
        slide_index: 0,
        acceptance_ids: ['A1.1'],
        passed: true,
      }),
    )
    await expect(skill.executeTool(call('verify_slides'))).resolves.toMatchObject({
      isError: true,
      output: expect.stringContaining('"error":"design_contract_production_incomplete"'),
    })
    expect(skill.buildContext?.()).toContain('"status":"producing"')
  })

  it('does not verify a complete contract when the host reports overlap', async () => {
    const proposals = createStructuredProposalController()
    const fake = adapter({
      verifySlides: vi.fn().mockResolvedValue({
        slideWidth: 960,
        slideHeight: 540,
        slides: [
          {
            slideId: 'slide-1',
            slideIndex: 0,
            shapes: [],
            shapesTruncated: false,
            overflows: [],
            overlaps: [{ shapeAId: '1', shapeBId: '2', overlapX: 10, overlapY: 10 }],
            overlapsTruncated: false,
          },
        ],
      }),
    })
    const skill = createPowerPointSkill({ adapter: fake, proposals })
    await skill.executeTool(call('plan_deck', { contract: modernContract() }))
    await skill.executeTool(
      call('edit_slide_text', { slide_index: 0, shape_id: '2', text: 'Hello' }),
    )
    await proposals.confirm(proposals.pending()!.id)
    await skill.executeTool(call('screenshot_slide', { slide_index: 0 }))
    await skill.executeTool(
      call('review_slide_screenshot', {
        slide_index: 0,
        acceptance_ids: ['A1.1'],
        passed: true,
      }),
    )

    const failed = await skill.executeTool(call('verify_slides'))
    expect(failed.isError).toBe(true)
    expect(JSON.parse(failed.output)).toMatchObject({
      error: 'design_contract_verification_failed',
      nextTool: 'list_slide_shapes',
      verification: {
        slides: [{ slideIndex: 0, overlaps: [{ shapeAId: '1', shapeBId: '2' }] }],
      },
    })
    // A page that already passed visual review must remain editable to repair geometry.
    const repair = await skill.executeTool(
      call('edit_slide_text', { slide_index: 0, shape_id: '2', text: 'Hello' }),
    )
    expect(repair.isError).not.toBe(true)
    await proposals.confirm(proposals.pending()!.id)
    expect(
      await skill.executeTool(
        call('review_slide_screenshot', {
          slide_index: 0,
          acceptance_ids: ['A1.1'],
          passed: true,
        }),
      ),
    ).toMatchObject({ isError: true, output: 'design_contract_screenshot_required' })
    expect(skill.buildContext?.()).toContain('"status":"producing"')
  })

  it('records a failed visual review as repair feedback without failing the tool', async () => {
    const proposals = createStructuredProposalController()
    const skill = createPowerPointSkill({ adapter: adapter(), proposals })
    await skill.executeTool(call('plan_deck', { contract: modernContract() }))
    await skill.executeTool(
      call('edit_slide_text', { slide_index: 0, shape_id: '2', text: 'Hello' }),
    )
    await proposals.confirm(proposals.pending()!.id)
    await skill.executeTool(call('screenshot_slide', { slide_index: 0 }))

    const review = await skill.executeTool(
      call('review_slide_screenshot', {
        slide_index: 0,
        acceptance_ids: ['A1.1'],
        passed: false,
        issues: ['Title contrast is too low'],
      }),
    )
    expect(review.isError).not.toBe(true)
    expect(JSON.parse(review.output)).toMatchObject({
      status: 'needs_repair',
      slide: 1,
      acceptanceIds: ['A1.1'],
      issues: ['Title contrast is too low'],
      nextTool: 'list_slide_shapes',
    })
    await skill.executeTool(call('screenshot_slide', { slide_index: 0 }))
    const unchanged = await skill.executeTool(
      call('review_slide_screenshot', {
        slide_index: 0,
        acceptance_ids: ['A1.1'],
        passed: true,
      }),
    )
    expect(unchanged.isError).not.toBe(true)
    expect(JSON.parse(unchanged.output)).toMatchObject({
      status: 'repair_required',
      slide: 1,
      nextTool: 'list_slide_shapes',
    })
    await expect(skill.executeTool(call('verify_slides'))).resolves.toMatchObject({
      isError: true,
      output: expect.stringContaining('"error":"design_contract_production_incomplete"'),
    })

    const repair = await skill.executeTool(
      call('edit_slide_text', { slide_index: 0, shape_id: '2', text: 'Hello' }),
    )
    expect(repair.isError).not.toBe(true)
    await proposals.confirm(proposals.pending()!.id)
    await expect(
      skill.executeTool(
        call('review_slide_screenshot', {
          slide_index: 0,
          acceptance_ids: ['A1.1'],
          passed: true,
        }),
      ),
    ).resolves.toMatchObject({
      isError: true,
      output: 'design_contract_screenshot_required',
    })
  })

  describe('visual review recovery', () => {
    async function setup() {
      const fake = adapter()
      const proposals = createStructuredProposalController()
      const skill = createPowerPointSkill({ adapter: fake, proposals })
      await skill.executeTool(call('plan_deck', { contract: modernContract() }))
      const edit = async () => {
        await skill.executeTool(
          call('edit_slide_text', { slide_index: 0, shape_id: '2', text: 'Hello' }),
        )
        return proposals.confirm(proposals.pending()!.id)
      }
      const screenshot = () => skill.executeTool(call('screenshot_slide', { slide_index: 0 }))
      const review = (passed: boolean) =>
        skill.executeTool(
          call('review_slide_screenshot', {
            slide_index: 0,
            acceptance_ids: ['A1.1'],
            passed,
            issues: passed ? [] : ['White title is unreadable'],
          }),
        )
      await edit()
      await screenshot()
      return { fake, skill, edit, screenshot, review }
    }

    it('invalidates an old visual pass when a fresh re-review reports a defect', async () => {
      const { skill, screenshot, review } = await setup()
      await review(true)
      await skill.executeTool(call('verify_slides'))
      expect(skill.buildContext?.()).toContain('"status":"verified"')
      await screenshot()
      expect(JSON.parse((await review(false)).output)).toMatchObject({ status: 'needs_repair' })
      expect(skill.buildContext?.()).toContain('"status":"producing"')
      expect(skill.buildContext?.()).toContain('"needsRepair":true')
      expect((await skill.executeTool(call('verify_slides'))).isError).toBe(true)
      await screenshot()
      expect(JSON.parse((await review(true)).output)).toMatchObject({ status: 'repair_required' })
    })

    it('accepts an applied-unverified repair only after a new screenshot and positive review', async () => {
      const { fake, skill, edit, screenshot, review } = await setup()
      await review(false)
      vi.mocked(fake.editSlideText).mockRejectedValueOnce(new Error('office_applied_unverified'))
      await edit()
      expect(await review(true)).toMatchObject({
        isError: true,
        output: 'design_contract_screenshot_required',
      })
      await screenshot()
      expect(JSON.parse((await review(true)).output)).toMatchObject({ status: 'passed' })
      expect((await skill.executeTool(call('verify_slides'))).isError).not.toBe(true)
    })

    it.each(['office_verify_failed', 'office_write_pending'])(
      'does not accept a %s write as an applied repair',
      async (code) => {
        const { fake, skill, edit, screenshot, review } = await setup()
        await review(false)
        vi.mocked(fake.editSlideText).mockRejectedValueOnce(new Error(code))
        if (code === 'office_verify_failed') await expect(edit()).rejects.toThrow(code)
        else await edit()
        await screenshot()
        expect(JSON.parse((await review(true)).output)).toMatchObject({ status: 'repair_required' })
        expect((await skill.executeTool(call('verify_slides'))).isError).toBe(true)
      },
    )

    it('directs native screenshot failure to same-page repair without clearing review gates', async () => {
      const { fake, skill, screenshot, review } = await setup()
      vi.mocked(fake.screenshotSlide).mockRejectedValueOnce(
        new Error('office_screenshot_unavailable'),
      )
      const result = await screenshot()
      expect(result.isError).toBe(true)
      expect(JSON.parse(result.output)).toMatchObject({
        error: 'office_read_failed',
        reason: 'office_screenshot_unavailable',
        slide_index: 0,
        nextTool: 'list_slide_shapes',
        repairAllowed: true,
        visualAvailableToModel: false,
      })
      expect(result.modelContent).toBeUndefined()
      expect(skill.validateImageMutation(0)).toBeUndefined()
      expect(skill.buildContext?.()).toContain('"screenshotUnavailable":true')
      expect(skill.reviewFinalResponse?.({ text: 'Done', mutated: true })).toContain(
        'list_slide_shapes',
      )
      expect(await review(true)).toMatchObject({
        isError: true,
        output: 'design_contract_screenshot_required',
      })
      expect((await skill.executeTool(call('verify_slides'))).isError).toBe(true)
      await screenshot()
      expect(skill.buildContext?.()).not.toContain('"screenshotUnavailable":true')
      expect(JSON.parse((await review(true)).output)).toMatchObject({ status: 'passed' })
    })

    it('keeps expansion blocked while allowing repair of a screenshot-failed batch page', async () => {
      const base = modernContract()
      const fake = adapter({
        screenshotSlide: vi.fn().mockRejectedValue(new Error('office_screenshot_unavailable')),
      })
      const proposals = createStructuredProposalController()
      const skill = createPowerPointSkill({ adapter: fake, proposals })
      await skill.executeTool(
        call('plan_deck', {
          contract: modernContract({
            prototypePages: [1, 2, 3],
            brief: { ...base.brief, pageCount: 4 },
            slides: [1, 2, 3, 4].map((number) => ({
              ...base.slides[0],
              number,
              acceptance: [{ id: `A${number}.1`, criterion: 'Readable' }],
            })),
          }),
        }),
      )
      for (const slide_index of [0, 1, 2]) {
        await skill.executeTool(
          call('edit_slide_text', { slide_index, shape_id: '2', text: 'Hello' }),
        )
        await proposals.confirm(proposals.pending()!.id)
      }
      await skill.executeTool(call('screenshot_slide', { slide_index: 2 }))
      expect(skill.validateImageMutation(2)).toBeUndefined()
      expect(JSON.parse(skill.validateImageMutation(3)!)).toMatchObject({
        error: 'design_contract_review_required',
        nextTool: 'list_slide_shapes',
        failedScreenshotSlideIndexes: [2],
      })
      expect((await skill.executeTool(call('verify_slides'))).isError).toBe(true)
    })
  })

  it('counts only the inserted page as produced when duplicating a slide', async () => {
    const base = modernContract()
    const first = (base.slides as Array<Record<string, unknown>>)[0]!
    const contract = modernContract({
      prototypePages: [1, 2],
      brief: { ...(base.brief as object), pageCount: 2 },
      slides: [
        first,
        {
          ...first,
          number: 2,
          title: 'Second page',
          acceptance: [{ id: 'A2.1', criterion: 'Readable' }],
        },
      ],
    })
    const proposals = createStructuredProposalController()
    const skill = createPowerPointSkill({
      adapter: adapter({
        duplicateSlide: vi.fn().mockResolvedValue({ slideId: 'copy' }),
        listSlideShapes: vi.fn().mockResolvedValue({
          slideId: 'copy',
          slideIndex: 1,
          shapes: [],
        }),
      }),
      proposals,
    })
    await skill.executeTool(call('plan_deck', { contract }))

    await skill.executeTool(call('duplicate_slide', { slide_index: 0 }))
    await proposals.confirm(proposals.pending()!.id)
    await skill.executeTool(call('screenshot_slide', { slide_index: 1 }))
    await skill.executeTool(
      call('review_slide_screenshot', {
        slide_index: 1,
        acceptance_ids: ['A2.1'],
        passed: true,
      }),
    )

    await expect(skill.executeTool(call('verify_slides'))).resolves.toMatchObject({
      isError: true,
      output: expect.stringContaining('"error":"design_contract_production_incomplete"'),
    })
  })

  it('keeps the agent in the screenshot and verification loop after each mutation batch', async () => {
    const skill = createPowerPointSkill({
      adapter: adapter(),
      proposals: createStructuredProposalController(),
    })

    expect(skill.repeatFinalResponseCorrection).toBe(true)

    await skill.executeTool(
      call('plan_deck', {
        core_hook: 'One clear story',
        style: 'Editorial',
        pages: [{ title: 'Cover', brief: 'Opening idea', image_queries: [] }],
      }),
    )
    await skill.executeTool(call('set_slide_background', { slide_index: 0, color: '#112233' }))

    expect(skill.reviewFinalResponse?.({ text: 'Done', mutated: true })).toContain(
      'screenshot_slide',
    )
    await skill.executeTool(call('screenshot_slide', { slide_index: 0 }))
    expect(skill.reviewFinalResponse?.({ text: 'Done', mutated: true })).toContain('verify_slides')
    await skill.executeTool(call('verify_slides'))
    expect(skill.reviewFinalResponse?.({ text: 'Done', mutated: true })).toBeUndefined()

    await skill.executeTool(call('set_slide_background', { slide_index: 0, color: '#223344' }))
    expect(skill.reviewFinalResponse?.({ text: 'Done', mutated: true })).toContain(
      'screenshot_slide',
    )
    await skill.executeTool(call('verify_slides'))
    await skill.executeTool(call('screenshot_slide', { slide_index: 0 }))
    expect(skill.reviewFinalResponse?.({ text: 'Done', mutated: true })).toContain('verify_slides')
  })

  it('does not let one screenshot certify multiple changed slides', async () => {
    const skill = createPowerPointSkill({
      adapter: adapter({
        getPresentationState: vi.fn().mockResolvedValue({
          host: 'powerpoint',
          slideCount: 2,
          selectedSlideIndexes: [0],
          capabilities: {},
        }),
      }),
      proposals: createStructuredProposalController(),
    })
    await skill.executeTool(call('get_presentation_state'))
    await skill.executeTool(call('set_slide_background', { slide_index: 0, color: '#112233' }))
    await skill.executeTool(call('set_slide_background', { slide_index: 1, color: '#223344' }))

    expect(skill.reviewFinalResponse?.({ text: 'Done', mutated: true })).toContain('(1, 2)')
    await skill.executeTool(call('screenshot_slide', { slide_index: 0 }))
    const correction = skill.reviewFinalResponse?.({ text: 'Done', mutated: true })
    expect(correction).toContain('(2)')
    expect(correction).not.toContain('(1, 2)')
    await skill.executeTool(call('screenshot_slide', { slide_index: 1 }))
    expect(skill.reviewFinalResponse?.({ text: 'Done', mutated: true })).toContain('verify_slides')
  })

  it('exposes deterministic presentation state and explicit zero-based slide contracts', async () => {
    const fake = adapter()
    const skill = createPowerPointSkill({
      adapter: fake,
      proposals: createStructuredProposalController(),
    })
    await expect(skill.executeTool(call('get_presentation_state'))).resolves.toMatchObject({
      output: expect.stringContaining('slideCount'),
      mutated: false,
    })
    expect(fake.getPresentationState).toHaveBeenCalledOnce()
    for (const tool of skill.tools) {
      const schema = tool.inputSchema as any
      if (schema.properties?.slide_index)
        expect(schema.properties.slide_index.description).toContain('index 0')
    }
    const officeJs = skill.tools.find((tool) => tool.name === 'execute_office_js')!
    const operations = (officeJs.inputSchema as any).properties.program.properties.operations.items
      .anyOf
    for (const operation of operations) {
      if (operation.properties?.slide_index)
        expect(operation.properties.slide_index.description).toContain('index 0')
    }
  })

  it('does not advertise the unreliable native master editor on PowerPoint for Mac', () => {
    const skill = createPowerPointSkill({
      adapter: adapter(),
      proposals: createStructuredProposalController(),
      platform: 'Mac',
      nativeMasterEditingSupported: true,
    })
    const names = skill.tools.map((tool) => tool.name)
    expect(names).not.toContain('edit_slide_master')
    expect(names).not.toContain('inspect_slide_masters')
    expect(skill.systemPrompt).toContain('build new decks with slide-level tools')
  })

  it('rejects malformed PowerPoint plans without creating a proposal', async () => {
    const proposals = createStructuredProposalController()
    const skill = createPowerPointSkill({ adapter: adapter(), proposals })
    await expect(
      skill.executeTool(call('plan_deck', { core_hook: '', style: 'plain', pages: [] })),
    ).resolves.toMatchObject({ isError: true, mutated: false, output: 'invalid_tool_input' })
    expect(proposals.pending()).toBeUndefined()
  })

  it('runs plan, confirmation-gated build, readback, and final verification in order', async () => {
    const order: string[] = []
    const fake = adapter({
      executeDeclarative: vi.fn().mockImplementation(async () => {
        order.push('write')
        return { createdShapeIds: ['created-1'] }
      }),
      listSlideShapes: vi
        .fn()
        .mockImplementationOnce(async () => {
          order.push('readback-shapes')
          return { slideId: 'slide-1', slideIndex: 0, shapes: [] }
        })
        .mockImplementation(async () => {
          order.push('readback')
          return {
            slideId: 'slide-1',
            slideIndex: 0,
            shapes: [
              {
                id: 'created-1',
                name: 'Deck title',
                type: 'TextBox',
                left: 60,
                top: 72,
                width: 840,
                height: 96,
              },
            ],
          }
        }),
      readSlideText: vi.fn().mockImplementation(async () => {
        order.push('readback-text')
        return {
          slideId: 'slide-1',
          shapeId: 'created-1',
          text: 'LLM onboarding',
          paragraphs: ['LLM onboarding'],
        }
      }),
      verifySlides: vi.fn().mockImplementation(async () => {
        order.push('verify')
        return { slideWidth: 960, slideHeight: 540, slides: [] }
      }),
    })
    const proposals = createStructuredProposalController()
    const skill = createPowerPointSkill({ adapter: fake, proposals })

    await skill.executeTool(
      call('plan_deck', {
        core_hook: 'Help new hires become productive with LLMs',
        style: 'Minimal blue onboarding deck',
        pages: [
          {
            title: 'LLM onboarding',
            type: 'cover',
            brief: 'Introduce the training goal',
            layout: 'hero_statement',
            image_queries: [],
          },
        ],
      }),
    )
    order.push('plan')
    await skill.executeTool(
      call('execute_office_js', {
        program: {
          version: 1,
          operations: [
            {
              op: 'add_text_box',
              slide_index: 0,
              name: 'Deck title',
              text: 'LLM onboarding',
              left: 60,
              top: 72,
              width: 840,
              height: 96,
            },
          ],
        },
      }),
    )
    order.push('proposal')
    expect(fake.executeDeclarative).not.toHaveBeenCalled()

    await proposals.confirm(proposals.pending()!.id)
    await skill.executeTool(call('verify_slides'))

    expect(order).toEqual([
      'plan',
      'verify',
      'proposal',
      'write',
      'readback-shapes',
      'readback',
      'readback-text',
      'verify',
    ])
  })

  it('exposes native master inspection and editing with exact schemas', () => {
    const skill = createPowerPointSkill({
      adapter: adapter(),
      proposals: createStructuredProposalController(),
    })
    expect(skill.tools.map((tool) => tool.name)).toEqual([
      'get_presentation_state',
      'inspect_slide_masters',
      'screenshot_slide',
      'review_slide_screenshot',
      'list_slide_shapes',
      'read_slide_text',
      'verify_slides',
      'set_slide_background',
      'ask_clarification',
      'plan_deck',
      'execute_office_js',
      'edit_slide_text',
      'edit_slide_xml',
      'edit_slide_chart',
      'edit_slide_master',
      'edit_slide_master_xml',
      'duplicate_slide',
    ])
    for (const tool of skill.tools) expect(tool.inputSchema.additionalProperties).toBe(false)
    expect(skill.tools.find((tool) => tool.name === 'edit_slide_text')?.inputSchema).toMatchObject({
      required: ['slide_index', 'shape_id', 'text'],
      properties: {
        slide_index: { type: 'integer', minimum: 0, maximum: 100000 },
        shape_id: { type: 'string', minLength: 1, maxLength: 256 },
        text: { type: 'string', maxLength: 12000 },
      },
    })
    for (const name of [
      'execute_office_js',
      'edit_slide_xml',
      'edit_slide_chart',
      'edit_slide_master',
      'edit_slide_master_xml',
    ]) {
      const schema = skill.tools.find((tool) => tool.name === name)?.inputSchema
      expect(schema?.properties).toHaveProperty('program')
      expect(schema?.properties).not.toHaveProperty('code')
      expect(schema?.required).toContain('program')
    }
    const executeSchema = skill.tools.find((tool) => tool.name === 'execute_office_js')
      ?.inputSchema as unknown as {
      properties: { program: { properties: { operations: { items: unknown } } } }
    }
    const programSchema = executeSchema.properties.program
    expect(programSchema.properties.operations.items).toHaveProperty('anyOf')
  })

  it('proposes one native master edit and verifies semantic readback', async () => {
    let state = await adapter().inspectSlideMasters()
    const fake = adapter({
      inspectSlideMasters: vi.fn().mockImplementation(() => Promise.resolve(state)),
      executeMasterOperations: vi.fn().mockImplementation(async () => {
        state = {
          masters: [
            {
              ...state.masters[0],
              background: { type: 'Solid', color: '#000000', transparency: 0 },
            },
          ],
        }
      }),
    })
    const proposals = createStructuredProposalController()
    const skill = createPowerPointSkill({ adapter: fake, proposals })
    await expect(
      skill.executeTool(
        call('edit_slide_master', {
          program: {
            version: 2,
            operations: [
              {
                op: 'set_master_background',
                master_id: 'master-1',
                fill: { type: 'solid', color: '#000000', transparency: 0 },
              },
            ],
          },
        }),
      ),
    ).resolves.toMatchObject({ mutated: false, summary: 'Proposed native PowerPoint master edit' })
    expect(proposals.pending()?.impact).toMatchObject({ count: 1, targets: ['master:master-1'] })
    await proposals.confirm(proposals.pending()!.id)
    expect(fake.executeMasterOperations).toHaveBeenCalledOnce()
  })

  it('recovers an already verified native master operation when a later operation fails', async () => {
    const original = await adapter().inspectSlideMasters()
    const state = structuredClone(original)
    let writes = 0
    const fake = adapter({
      inspectSlideMasters: vi
        .fn()
        .mockImplementation(() => Promise.resolve(structuredClone(state))),
      executeMasterOperations: vi.fn().mockImplementation(async (operations) => {
        writes += 1
        const operation = operations[0]
        if (writes === 2) throw new Error('office_write_failed')
        if (operation.op === 'set_master_background' && operation.fill.type === 'solid')
          state.masters[0]!.background = {
            type: 'Solid',
            color: operation.fill.color,
            transparency: operation.fill.transparency,
          }
        if (operation.op === 'set_master_theme_color')
          state.masters[0]!.themeColors[operation.theme_color] = operation.color
      }),
    })
    const proposals = createStructuredProposalController()
    const skill = createPowerPointSkill({ adapter: fake, proposals })
    await skill.executeTool(
      call('edit_slide_master', {
        program: {
          version: 2,
          operations: [
            {
              op: 'set_master_background',
              master_id: 'master-1',
              fill: { type: 'solid', color: '#000000', transparency: 0 },
            },
            {
              op: 'set_master_theme_color',
              master_id: 'master-1',
              theme_color: 'Light1',
              color: '#EEEEEE',
            },
          ],
        },
      }),
    )
    await expect(proposals.confirm(proposals.pending()!.id)).rejects.toThrow('office_write_failed')
    expect(state).toEqual(original)
    expect(fake.executeMasterOperations).toHaveBeenCalledTimes(3)
  })

  it('does not advertise or execute master package edits on PowerPoint for Mac', async () => {
    const fake = adapter()
    const skill = createPowerPointSkill({
      adapter: fake,
      proposals: createStructuredProposalController(),
      platform: 'Mac',
    })

    expect(skill.tools.map((tool) => tool.name)).not.toContain('edit_slide_master')
    expect(skill.tools.map((tool) => tool.name)).not.toContain('inspect_slide_masters')
    expect(skill.tools.map((tool) => tool.name)).not.toContain('edit_slide_master_xml')
    expect(skill.systemPrompt).toContain('build new decks with slide-level tools')
    await expect(
      skill.executeTool(
        call('edit_slide_master_xml', {
          program: {
            version: 1,
            operations: [
              {
                op: 'replace_xml',
                path: 'ppt/slideMasters/slideMaster1.xml',
                xml: '<p:sldMaster xmlns:p="urn:p"/>',
              },
            ],
          },
        }),
      ),
    ).resolves.toMatchObject({ isError: true, output: 'office_api_unsupported' })
    expect(fake.exportSlidePackage).not.toHaveBeenCalled()
    expect(fake.replaceSlidePackage).not.toHaveBeenCalled()
  })

  it('normalizes reads, image display, and rejects unknown fields', async () => {
    const fake = adapter()
    const skill = createPowerPointSkill({
      adapter: fake,
      proposals: createStructuredProposalController(),
    })
    await expect(
      skill.executeTool(call('list_slide_shapes', { slide_index: 0 })),
    ).resolves.toMatchObject({
      mutated: false,
      output: expect.stringContaining('"id":"2"'),
    })
    await expect(
      skill.executeTool(call('read_slide_text', { slide_index: 0, shape_id: '2' })),
    ).resolves.toMatchObject({
      mutated: false,
      output: expect.stringContaining('Hello'),
    })
    await expect(
      skill.executeTool(call('screenshot_slide', { slide_index: 0 })),
    ).resolves.toMatchObject({
      mutated: false,
      output: expect.stringContaining('"visualAvailableToModel":true'),
      modelContent: [{ type: 'image', image: { mime: 'image/png', base64: png } }],
      display: { kind: 'images', items: [{ url: `data:image/png;base64,${png}` }] },
    })
    await expect(skill.executeTool(call('verify_slides', { nope: true }))).resolves.toMatchObject({
      output: 'invalid_tool_input',
      isError: true,
    })
  })

  it('gates text edits behind immutable stale-checked proposals and verifies after confirmation', async () => {
    const fake = adapter({
      readSlideText: vi
        .fn()
        .mockResolvedValueOnce({
          slideId: 'slide-1',
          shapeId: '2',
          text: 'Hello',
          paragraphs: ['Hello'],
        })
        .mockResolvedValueOnce({
          slideId: 'slide-1',
          shapeId: '2',
          text: 'Hello',
          paragraphs: ['Hello'],
        })
        .mockResolvedValue({
          slideId: 'slide-1',
          shapeId: '2',
          text: 'New',
          paragraphs: ['New'],
        }),
    })
    const proposals = createStructuredProposalController()
    const skill = createPowerPointSkill({ adapter: fake, proposals })
    const proposed = await skill.executeTool(
      call('edit_slide_text', {
        slide_index: 0,
        shape_id: '2',
        text: 'New',
        explanation: 'Update title',
      }),
    )
    expect(proposed).toMatchObject({ mutated: false, summary: 'Proposed PowerPoint text edit' })
    const pending = proposals.pending()!
    expect(pending).toMatchObject({
      toolName: 'edit_slide_text',
      preview: { shapeId: '2', before: 'Hello', after: 'New' },
      impact: { host: 'powerpoint', targets: ['slide-1/2'], count: 1 },
    })
    await proposals.confirm(pending.id)
    expect(fake.editSlideText).toHaveBeenCalledWith(0, '2', 'New', expect.any(AbortSignal))
    expect(fake.snapshotSlide).not.toHaveBeenCalled()
    expect(fake.verifySlides).toHaveBeenCalledOnce()
  })

  it('refuses stale or cancelled writes before mutation', async () => {
    const fake = adapter({
      listSlideShapes: vi.fn().mockImplementation((index: number) =>
        Promise.resolve({
          slideId: index === 1 ? 'slide-copy' : 'slide-1',
          slideIndex: index,
          shapes: [],
        }),
      ),
    })
    const proposals = createStructuredProposalController()
    const skill = createPowerPointSkill({ adapter: fake, proposals })
    await skill.executeTool(call('duplicate_slide', { slide_index: 0 }))
    ;(fake.snapshotSlide as ReturnType<typeof vi.fn>).mockResolvedValueOnce({
      slideId: 'slide-1',
      fingerprint: 'changed',
    })
    await expect(proposals.confirm(proposals.pending()!.id)).rejects.toThrow('proposal_stale')
    expect(fake.duplicateSlide).not.toHaveBeenCalled()

    const controller = new AbortController()
    controller.abort()
    await expect(
      skill.executeTool(
        call('edit_slide_text', { slide_index: 0, shape_id: '2', text: 'x' }),
        controller.signal,
      ),
    ).resolves.toMatchObject({
      output: 'cancelled',
      isError: true,
    })
  })

  it('rejects JavaScript syntax and unknown declarative authority without proposals', async () => {
    const proposals = createStructuredProposalController()
    const skill = createPowerPointSkill({ adapter: adapter(), proposals })
    for (const input of [
      { code: 'return context.presentation' },
      { code: '{"version":1,"operations":[{"op":"fetch","url":"https://x"}]}' },
    ]) {
      await expect(skill.executeTool(call('execute_office_js', input))).resolves.toMatchObject({
        output: 'invalid_tool_input',
        isError: true,
        mutated: false,
      })
      expect(proposals.pending()).toBeUndefined()
    }
  })

  it('accepts direct structured programs without JSON string double encoding', async () => {
    const fake = adapter()
    const proposals = createStructuredProposalController()
    const skill = createPowerPointSkill({ adapter: fake, proposals })
    const program = {
      version: 1,
      operations: [{ op: 'set_shape_text', slide_index: 0, shape_id: '2', text: 'Structured' }],
    }

    await expect(skill.executeTool(call('execute_office_js', { program }))).resolves.toMatchObject({
      mutated: false,
      output: expect.stringContaining('set_shape_text'),
    })
    expect(proposals.pending()?.preview).toEqual(program)
  })

  it('rejects stale shape ids before creating a PowerPoint write proposal', async () => {
    const fake = adapter({
      listSlideShapes: vi.fn().mockResolvedValue({
        slideId: 'slide-1',
        slideIndex: 0,
        shapes: [],
      }),
    })
    const proposals = createStructuredProposalController()
    const skill = createPowerPointSkill({ adapter: fake, proposals })

    await expect(
      skill.executeTool(
        call('execute_office_js', {
          program: {
            version: 1,
            operations: [
              { op: 'set_shape_text_style', slide_index: 0, shape_id: 'deleted-shape', bold: true },
            ],
          },
        }),
      ),
    ).resolves.toMatchObject({ output: 'invalid_tool_input', isError: true, mutated: false })
    expect(proposals.pending()).toBeUndefined()
    expect(fake.executeDeclarative).not.toHaveBeenCalled()
  })

  it('reports a content-free program location for invalid operation fields', async () => {
    const skill = createPowerPointSkill({
      adapter: adapter(),
      proposals: createStructuredProposalController(),
    })
    await expect(
      skill.executeTool(
        call('execute_office_js', {
          program: { version: 1, operations: [{ op: 'add_text_box', slide_index: 0 }] },
        }),
      ),
    ).resolves.toMatchObject({
      output: 'invalid_tool_input',
      isError: true,
      diagnosticError: {
        code: 'InvalidToolInput',
        debugInfo: { errorLocation: 'program.operations' },
      },
    })
  })

  it('reports a content-free program location for malformed streamed tool JSON', async () => {
    const skill = createPowerPointSkill({
      adapter: adapter(),
      proposals: createStructuredProposalController(),
    })
    await expect(
      skill.executeTool({
        id: 'call-1',
        name: 'execute_office_js',
        input: {},
        inputError: 'raw malformed JSON must not be retained',
      }),
    ).resolves.toMatchObject({
      output: 'invalid_tool_input',
      isError: true,
      diagnosticError: {
        code: 'InvalidToolInput',
        debugInfo: { errorLocation: 'program' },
      },
    })
  })

  it('accepts direct structured XML programs for slide and master edits', async () => {
    const zip = new JSZip()
    zip.file('ppt/slides/slide1.xml', '<p:sld xmlns:p="urn:p"/>')
    zip.file('ppt/slideMasters/slideMaster1.xml', '<p:sldMaster xmlns:p="urn:p"/>')
    const base64 = await zip.generateAsync({ type: 'base64' })
    const fake = adapter({
      exportSlidePackage: vi.fn().mockResolvedValue({
        slideId: 's1',
        base64,
        fingerprint: 'stable',
      }),
    })
    for (const [name, input] of [
      [
        'edit_slide_xml',
        {
          slide_index: 0,
          program: {
            version: 1,
            operations: [
              {
                op: 'replace_xml',
                path: 'ppt/slides/slide1.xml',
                xml: '<p:sld xmlns:p="urn:p"><p:cSld/></p:sld>',
              },
            ],
          },
        },
      ],
      [
        'edit_slide_master_xml',
        {
          program: {
            version: 1,
            operations: [
              {
                op: 'replace_xml',
                path: 'ppt/slideMasters/slideMaster1.xml',
                xml: '<p:sldMaster xmlns:p="urn:p"><p:cSld/></p:sldMaster>',
              },
            ],
          },
        },
      ],
    ] as const) {
      const controller = createStructuredProposalController()
      const scoped = createPowerPointSkill({ adapter: fake, proposals: controller })
      await expect(scoped.executeTool(call(name, input))).resolves.toMatchObject({
        mutated: false,
        summary: expect.stringContaining('Proposed'),
      })
      controller.reject()
    }
  })

  it('proposes and semantically verifies bounded XML, chart, and master package edits', async () => {
    const zip = new JSZip()
    zip.file('ppt/slides/slide1.xml', '<p:sld xmlns:p="urn:p"/>')
    zip.file('ppt/charts/chart1.xml', '<c:chart xmlns:c="urn:c"/>')
    zip.file('ppt/slideMasters/slideMaster1.xml', '<p:sldMaster xmlns:p="urn:p"/>')
    let current = await zip.generateAsync({ type: 'base64' })
    const fake = adapter({
      exportSlidePackage: vi.fn().mockImplementation(() =>
        Promise.resolve({
          slideId: 's1',
          base64: current,
          fingerprint: `${current.length}:${current.slice(-8)}`,
        }),
      ),
      replaceSlidePackage: vi.fn().mockImplementation((_index, base64) => {
        current = base64
        return Promise.resolve({ slideId: 's2' })
      }),
    })
    for (const [name, input] of [
      [
        'edit_slide_xml',
        {
          slide_index: 0,
          code: JSON.stringify({
            version: 1,
            operations: [
              {
                op: 'replace_xml',
                path: 'ppt/slides/slide1.xml',
                xml: '<p:sld xmlns:p="urn:p"><p:cSld/></p:sld>',
              },
            ],
          }),
        },
      ],
      [
        'edit_slide_chart',
        {
          slide_index: 0,
          code: JSON.stringify({
            version: 1,
            operations: [
              {
                op: 'replace_xml',
                path: 'ppt/charts/chart1.xml',
                xml: '<c:chart xmlns:c="urn:c"><c:title/></c:chart>',
              },
            ],
          }),
        },
      ],
      [
        'edit_slide_master_xml',
        {
          code: JSON.stringify({
            version: 1,
            operations: [
              {
                op: 'replace_xml',
                path: 'ppt/slideMasters/slideMaster1.xml',
                xml: '<p:sldMaster xmlns:p="urn:p"><p:cSld/></p:sldMaster>',
              },
            ],
          }),
        },
      ],
    ] as const) {
      const proposals = createStructuredProposalController()
      const skill = createPowerPointSkill({ adapter: fake, proposals })
      await expect(skill.executeTool(call(name, input))).resolves.toMatchObject({
        mutated: false,
        summary: expect.stringContaining('Proposed'),
      })
      await proposals.confirm(proposals.pending()!.id)
    }
    expect(fake.replaceSlidePackage).toHaveBeenCalledTimes(3)
  })

  it('validates master edits from the targeted XML instead of volatile package bytes', async () => {
    const first = new JSZip()
    first.file('ppt/slideMasters/slideMaster1.xml', '<p:sldMaster xmlns:p="urn:p"/>')
    first.file('docProps/core.xml', '<core modified="one"/>')
    const second = new JSZip()
    second.file('ppt/slideMasters/slideMaster1.xml', '<p:sldMaster xmlns:p="urn:p"/>')
    second.file('docProps/core.xml', '<core modified="two"/>')
    let current = await first.generateAsync({ type: 'base64' })
    const confirmSnapshot = await second.generateAsync({ type: 'base64' })
    let exports = 0
    const fake = adapter({
      exportSlidePackage: vi.fn().mockImplementation(() => {
        exports += 1
        const base64 = exports === 2 || exports === 3 ? confirmSnapshot : current
        return Promise.resolve({ slideId: 's1', base64, fingerprint: `volatile-${exports}` })
      }),
      replaceSlidePackage: vi.fn().mockImplementation((_index, base64) => {
        current = base64
        return Promise.resolve({ slideId: 's1' })
      }),
    })
    const proposals = createStructuredProposalController()
    const skill = createPowerPointSkill({ adapter: fake, proposals })

    await skill.executeTool(
      call('edit_slide_master_xml', {
        program: {
          version: 1,
          operations: [
            {
              op: 'replace_xml',
              path: 'ppt/slideMasters/slideMaster1.xml',
              xml: '<p:sldMaster xmlns:p="urn:p"><p:cSld/></p:sldMaster>',
            },
          ],
        },
      }),
    )
    await expect(proposals.confirm(proposals.pending()!.id)).resolves.toBeUndefined()
    expect(fake.replaceSlidePackage).toHaveBeenCalledOnce()
    const appliedBase64 = vi.mocked(fake.replaceSlidePackage).mock.calls[0]?.[1]
    const applied = await JSZip.loadAsync(appliedBase64!, { base64: true })
    await expect(applied.file('docProps/core.xml')?.async('string')).resolves.toContain('two')
  })

  it('does not overwrite a target XML change between validation and execution', async () => {
    const original = new JSZip()
    original.file('ppt/slideMasters/slideMaster1.xml', '<p:sldMaster xmlns:p="urn:p"/>')
    const changed = new JSZip()
    changed.file(
      'ppt/slideMasters/slideMaster1.xml',
      '<p:sldMaster xmlns:p="urn:p"><p:changed-by-user/></p:sldMaster>',
    )
    const originalBase64 = await original.generateAsync({ type: 'base64' })
    const changedBase64 = await changed.generateAsync({ type: 'base64' })
    let exports = 0
    const fake = adapter({
      exportSlidePackage: vi.fn().mockImplementation(() => {
        exports += 1
        return Promise.resolve({
          slideId: 's1',
          base64: exports < 3 ? originalBase64 : changedBase64,
          fingerprint: `volatile-${exports}`,
        })
      }),
      replaceSlidePackage: vi.fn(),
    })
    const proposals = createStructuredProposalController()
    const skill = createPowerPointSkill({ adapter: fake, proposals })
    await skill.executeTool(
      call('edit_slide_master_xml', {
        program: {
          version: 1,
          operations: [
            {
              op: 'replace_xml',
              path: 'ppt/slideMasters/slideMaster1.xml',
              xml: '<p:sldMaster xmlns:p="urn:p"><p:cSld/></p:sldMaster>',
            },
          ],
        },
      }),
    )

    await expect(proposals.confirm(proposals.pending()!.id)).rejects.toThrow('proposal_stale')
    expect(fake.replaceSlidePackage).not.toHaveBeenCalled()
  })

  it('executes only confirmed declarative PowerPoint operations and verifies text', async () => {
    const fake = adapter({
      exportSlidePackage: vi
        .fn()
        .mockResolvedValue({ slideId: 's1', base64: 'ppt', fingerprint: 'same' }),
      executeDeclarative: vi.fn().mockResolvedValue({ createdShapeIds: [] }),
      readSlideText: vi
        .fn()
        .mockResolvedValue({ slideId: 's1', shapeId: '2', text: 'New', paragraphs: ['New'] }),
    })
    const proposals = createStructuredProposalController()
    const skill = createPowerPointSkill({ adapter: fake, proposals })
    const code =
      '{"version":1,"operations":[{"op":"set_shape_text","slide_index":0,"shape_id":"2","text":"New"}]}'
    await skill.executeTool(call('execute_office_js', { code }))
    expect(fake.executeDeclarative).not.toHaveBeenCalled()
    await proposals.confirm(proposals.pending()!.id)
    expect(fake.executeDeclarative).toHaveBeenCalledWith(
      [{ op: 'set_shape_text', slide_index: 0, shape_id: '2', text: 'New' }],
      expect.any(AbortSignal),
    )
    expect(fake.snapshotSlide).toHaveBeenCalledTimes(2)
    expect(fake.exportSlidePackage).not.toHaveBeenCalled()
  })

  it('confirmation-gates bounded text style and verifies exact color readback', async () => {
    let color = '#000000'
    const fake = adapter({
      executeDeclarative: vi.fn().mockImplementation(async (operations) => {
        for (const operation of operations)
          if (operation.op === 'set_shape_text_style' && operation.color) color = operation.color
        return { createdShapeIds: [] }
      }),
      readShapeTextStyle: vi.fn().mockImplementation(async () => ({ color })),
    })
    const proposals = createStructuredProposalController()
    const skill = createPowerPointSkill({ adapter: fake, proposals })
    await skill.executeTool(
      call('execute_office_js', {
        program: {
          version: 1,
          operations: [
            { op: 'set_shape_text_style', slide_index: 0, shape_id: '2', color: '#111111' },
            { op: 'set_shape_text_style', slide_index: 0, shape_id: '2', color: '#2457A7' },
          ],
        },
      }),
    )
    expect(color).toBe('#000000')
    await proposals.confirm(proposals.pending()!.id)
    expect(color).toBe('#2457A7')
    expect(fake.readShapeTextStyle).toHaveBeenCalled()
  })

  it('accepts semantically equivalent Mac PowerPoint font readback', async () => {
    const fake = adapter({
      executeDeclarative: vi.fn().mockResolvedValue({ createdShapeIds: [] }),
      readShapeTextStyle: vi.fn().mockResolvedValue({
        fontFamily: '  APTOS   DISPLAY ',
        fontSize: 23.999,
      }),
    })
    const proposals = createStructuredProposalController()
    const skill = createPowerPointSkill({ adapter: fake, proposals })
    await skill.executeTool(
      call('execute_office_js', {
        program: {
          version: 1,
          operations: [
            {
              op: 'set_shape_text_style',
              slide_index: 0,
              shape_id: '2',
              fontFamily: 'Aptos Display',
              fontSize: 24,
            },
          ],
        },
      }),
    )
    await expect(proposals.confirm(proposals.pending()!.id)).resolves.toBeUndefined()
  })

  it.each([{ color: 'red' }, { fontSize: 0 }, { fontFamily: '' }, { bold: 'yes' }])(
    'rejects malformed declarative text style without a proposal: %j',
    async (style) => {
      const proposals = createStructuredProposalController()
      const skill = createPowerPointSkill({ adapter: adapter(), proposals })
      await expect(
        skill.executeTool(
          call('execute_office_js', {
            program: {
              version: 1,
              operations: [{ op: 'set_shape_text_style', slide_index: 0, shape_id: '2', ...style }],
            },
          }),
        ),
      ).resolves.toMatchObject({ isError: true, mutated: false })
      expect(proposals.pending()).toBeUndefined()
    },
  )

  it('accepts strict declarative geometry, text-box creation, and shape deletion families', async () => {
    const fake = adapter({
      listSlideShapes: vi.fn().mockResolvedValue({
        slideId: 'slide-1',
        slideIndex: 0,
        shapes: [
          { id: '2', name: 'Title', type: 'TextBox', left: 10, top: 20, width: 200, height: 40 },
          { id: '9', name: 'Placeholder', type: 'TextBox', left: 0, top: 0, width: 20, height: 20 },
        ],
      }),
      exportSlidePackage: vi.fn().mockResolvedValue({
        slideId: 's1',
        base64: 'ppt',
        fingerprint: 'same',
      }),
    })
    const proposals = createStructuredProposalController()
    const skill = createPowerPointSkill({ adapter: fake, proposals })
    const code = JSON.stringify({
      version: 1,
      operations: [
        {
          op: 'set_shape_geometry',
          slide_index: 0,
          shape_id: '2',
          left: 1,
          top: 2,
          width: 3,
          height: 4,
        },
        {
          op: 'add_text_box',
          slide_index: 0,
          name: 'Agent box',
          text: 'Hi',
          left: 5,
          top: 6,
          width: 70,
          height: 20,
        },
        { op: 'delete_shape', slide_index: 0, shape_id: '9' },
      ],
    })
    await expect(skill.executeTool(call('execute_office_js', { code }))).resolves.toMatchObject({
      mutated: false,
      output: expect.stringContaining('set_shape_geometry'),
    })
    expect(proposals.pending()?.impact.count).toBe(3)
    proposals.reject()
    await expect(
      skill.executeTool(
        call('execute_office_js', {
          code: JSON.stringify({
            version: 1,
            operations: [
              {
                op: 'set_shape_geometry',
                slide_index: 0,
                shape_id: '2',
                left: 1,
                top: 2,
                width: -1,
                height: 4,
              },
            ],
          }),
        }),
      ),
    ).resolves.toMatchObject({ output: 'invalid_tool_input', isError: true })
  })

  it('accepts host-normalized PowerPoint geometry when verifying a created text box', async () => {
    const fake = adapter({
      executeDeclarative: vi.fn().mockResolvedValue({ createdShapeIds: ['4'] }),
      listSlideShapes: vi
        .fn()
        .mockResolvedValueOnce({ slideId: 's1', slideIndex: 0, shapes: [] })
        .mockResolvedValue({
          slideId: 's1',
          slideIndex: 0,
          shapes: [
            {
              id: '4',
              name: 'Status',
              type: 'TextBox',
              left: 300.00003,
              top: 449.99997,
              width: 360.00003,
              height: 50.00003,
            },
          ],
        }),
      readSlideText: vi.fn().mockResolvedValue({
        slideId: 's1',
        shapeId: '4',
        text: 'PASS',
        paragraphs: ['PASS'],
      }),
    })
    const proposals = createStructuredProposalController()
    const skill = createPowerPointSkill({ adapter: fake, proposals })
    const code = JSON.stringify({
      version: 1,
      operations: [
        {
          op: 'add_text_box',
          slide_index: 0,
          name: 'Status',
          text: 'PASS',
          left: 300,
          top: 450,
          width: 360,
          height: 50,
        },
      ],
    })

    await skill.executeTool(call('execute_office_js', { code }))
    await expect(proposals.confirm(proposals.pending()!.id)).resolves.toBeUndefined()
  })

  it('accepts PowerPoint paragraph normalization when verifying created text', async () => {
    const fake = adapter({
      executeDeclarative: vi.fn().mockResolvedValue({ createdShapeIds: ['4'] }),
      listSlideShapes: vi
        .fn()
        .mockResolvedValueOnce({ slideId: 's1', slideIndex: 0, shapes: [] })
        .mockResolvedValue({
          slideId: 's1',
          slideIndex: 0,
          shapes: [
            {
              id: '4',
              name: 'Summary',
              type: 'TextBox',
              left: 100,
              top: 100,
              width: 400,
              height: 100,
            },
          ],
        }),
      readSlideText: vi.fn().mockResolvedValue({
        slideId: 's1',
        shapeId: '4',
        text: 'First line\rSecond line\vThird line',
        paragraphs: ['First line', 'Second line', 'Third line'],
      }),
    })
    const proposals = createStructuredProposalController()
    const skill = createPowerPointSkill({ adapter: fake, proposals })

    await skill.executeTool(
      call('execute_office_js', {
        program: {
          version: 1,
          operations: [
            {
              op: 'add_text_box',
              slide_index: 0,
              name: 'Summary',
              text: 'First line\nSecond line\nThird line',
              left: 100,
              top: 100,
              width: 400,
              height: 100,
            },
          ],
        },
      }),
    )

    await expect(proposals.confirm(proposals.pending()!.id)).resolves.toBeUndefined()
  })

  it('waits for delayed PowerPoint text readback before rejecting an applied edit', async () => {
    const fake = adapter({
      readSlideText: vi
        .fn()
        .mockResolvedValueOnce({
          slideId: 's1',
          shapeId: '2',
          text: 'Old',
          paragraphs: ['Old'],
        })
        .mockResolvedValueOnce({
          slideId: 's1',
          shapeId: '2',
          text: 'Old',
          paragraphs: ['Old'],
        })
        .mockResolvedValueOnce({
          slideId: 's1',
          shapeId: '2',
          text: 'Old',
          paragraphs: ['Old'],
        })
        .mockResolvedValue({
          slideId: 's1',
          shapeId: '2',
          text: 'New',
          paragraphs: ['New'],
        }),
    })
    const proposals = createStructuredProposalController()
    const skill = createPowerPointSkill({ adapter: fake, proposals })

    await skill.executeTool(call('edit_slide_text', { slide_index: 0, shape_id: '2', text: 'New' }))

    await expect(proposals.confirm(proposals.pending()!.id)).resolves.toBeUndefined()
    expect(fake.readSlideText).toHaveBeenCalledTimes(4)
  })

  it('waits for delayed geometry, delete, and duplicate readback', async () => {
    const beforeShape = {
      id: '2',
      name: 'Title',
      type: 'TextBox',
      left: 10,
      top: 20,
      width: 200,
      height: 40,
    }
    const afterShape = { ...beforeShape, left: 30 }
    const deletedShape = { ...beforeShape, id: '9', name: 'Remove me' }
    const fake = adapter({
      executeDeclarative: vi.fn().mockResolvedValue({ createdShapeIds: [] }),
      listSlideShapes: vi
        .fn()
        .mockResolvedValueOnce({
          slideId: 's1',
          slideIndex: 0,
          shapes: [beforeShape, deletedShape],
        })
        .mockResolvedValueOnce({
          slideId: 's1',
          slideIndex: 0,
          shapes: [afterShape, deletedShape],
        })
        .mockResolvedValueOnce({
          slideId: 's1',
          slideIndex: 0,
          shapes: [afterShape, deletedShape],
        })
        .mockResolvedValueOnce({ slideId: 's1', slideIndex: 0, shapes: [afterShape] }),
    })
    const proposals = createStructuredProposalController()
    const skill = createPowerPointSkill({ adapter: fake, proposals })

    await skill.executeTool(
      call('execute_office_js', {
        code: JSON.stringify({
          version: 1,
          operations: [
            {
              op: 'set_shape_geometry',
              slide_index: 0,
              shape_id: '2',
              left: 30,
              top: 20,
              width: 200,
              height: 40,
            },
            { op: 'delete_shape', slide_index: 0, shape_id: '9' },
          ],
        }),
      }),
    )

    await expect(proposals.confirm(proposals.pending()!.id)).resolves.toBeUndefined()
  })

  it('waits for a delayed duplicate-slide collection readback', async () => {
    const fake = adapter({
      duplicateSlide: vi.fn().mockResolvedValue({ slideId: 'copy' }),
      listSlideShapes: vi
        .fn()
        .mockResolvedValueOnce({ slideId: 'slide-1', slideIndex: 1, shapes: [] })
        .mockResolvedValueOnce({ slideId: 'slide-1', slideIndex: 1, shapes: [] })
        .mockResolvedValue({ slideId: 'copy', slideIndex: 1, shapes: [] }),
    })
    const proposals = createStructuredProposalController()
    const skill = createPowerPointSkill({ adapter: fake, proposals })

    await skill.executeTool(call('duplicate_slide', { slide_index: 0 }))
    await expect(proposals.confirm(proposals.pending()!.id)).resolves.toBeUndefined()
    expect(fake.listSlideShapes).toHaveBeenCalledTimes(3)
  })

  it.each(['duplicate_slide', 'execute_office_js'])(
    'allows successive duplicates after reading the slide count through %s',
    async (toolName) => {
      const slideIds = ['slide-1']
      const duplicateSlide = vi.fn(async (index: number) => {
        const slideId = `copy-${slideIds.length}`
        slideIds.splice(index + 1, 0, slideId)
        return { slideId }
      })
      const fake = adapter({
        duplicateSlide,
        executeDeclarative: vi.fn(async (operations) => ({
          createdShapeIds: [],
          insertedSlideId: (await duplicateSlide(operations[0]!.slide_index)).slideId,
        })),
        snapshotSlide: vi.fn(async (index) => ({
          slideId: slideIds[index]!,
          fingerprint: slideIds[index]!,
        })),
        listSlideShapes: vi.fn(async (index) => ({
          slideId: slideIds[index]!,
          slideIndex: index,
          shapes: [],
        })),
      })
      const proposals = createStructuredProposalController()
      const skill = createPowerPointSkill({ adapter: fake, proposals })
      await skill.executeTool(call('get_presentation_state'))

      for (const slideIndex of [0, 1]) {
        const result = await skill.executeTool(
          call(
            toolName,
            toolName === 'duplicate_slide'
              ? { slide_index: slideIndex }
              : {
                  program: {
                    version: 1,
                    operations: [{ op: 'duplicate_slide', slide_index: slideIndex }],
                  },
                },
          ),
        )
        expect(result.isError).not.toBe(true)
        await proposals.confirm(proposals.pending()!.id)
      }

      expect(slideIds).toEqual(['slide-1', 'copy-1', 'copy-2'])
      expect(
        await skill.executeTool(
          call('edit_slide_text', { slide_index: 2, shape_id: '2', text: 'Hello' }),
        ),
      ).not.toHaveProperty('isError', true)
      await expect(
        skill.executeTool(call('duplicate_slide', { slide_index: 3 })),
      ).resolves.toMatchObject({ isError: true, output: 'invalid_tool_input' })
    },
  )

  it('keeps prototype and review gates while duplicating newly created slides', async () => {
    const base = modernContract()
    const contract = modernContract({
      prototypePages: [1, 2, 3],
      brief: { ...base.brief, pageCount: 4 },
      slides: [1, 2, 3, 4].map((number) => ({
        ...base.slides[0],
        number,
        acceptance: [{ id: `A${number}.1`, criterion: 'Readable' }],
      })),
    })
    const proposals = createStructuredProposalController()
    const skill = createPowerPointSkill({
      adapter: adapter({
        listSlideShapes: vi
          .fn()
          .mockResolvedValue({ slideId: 'slide-copy', slideIndex: 1, shapes: [] }),
      }),
      proposals,
    })
    await skill.executeTool(call('get_presentation_state'))
    expect(await skill.executeTool(call('plan_deck', { contract }))).not.toHaveProperty(
      'isError',
      true,
    )
    await skill.executeTool(
      call('edit_slide_text', { slide_index: 0, shape_id: '2', text: 'Hello' }),
    )
    await proposals.confirm(proposals.pending()!.id)

    for (const slideIndex of [0, 1]) {
      expect(
        await skill.executeTool(call('duplicate_slide', { slide_index: slideIndex })),
      ).not.toHaveProperty('isError', true)
      await proposals.confirm(proposals.pending()!.id)
    }

    await expect(
      skill.executeTool(call('duplicate_slide', { slide_index: 2 })),
    ).resolves.toMatchObject({
      isError: true,
      output: expect.stringContaining('"error":"design_contract_review_required"'),
    })
    expect(proposals.pending()).toBeUndefined()
    expect(skill.buildContext?.()).toContain('"status":"producing"')
  })

  describe('non-contiguous prototype scaffolding', () => {
    async function setup(
      options: {
        prototypePages?: number[]
        slideCount?: number
        pageCount?: number
        status?: string
        readState?: boolean
      } = {},
    ) {
      const base = modernContract()
      const pageCount = options.pageCount ?? 4
      const contract = modernContract({
        status: options.status ?? 'ready',
        prototypePages: options.prototypePages ?? [1, 3, 4],
        brief: { ...base.brief, pageCount },
        slides: Array.from({ length: pageCount }, (_, index) => ({
          ...base.slides[0],
          number: index + 1,
          acceptance: [{ id: `A${index + 1}.1`, criterion: 'Readable' }],
        })),
      })
      const slideIds = Array.from(
        { length: options.slideCount ?? 1 },
        (_, index) => `slide-${index}`,
      )
      const fake = adapter({
        getPresentationState: vi.fn(async () => ({
          slideCount: slideIds.length,
          selectedSlideIndexes: [0],
          api: { v12: true, v14: true, v15: true, v18: true, v110: true },
        })),
        duplicateSlide: vi.fn(async (index) => {
          const slideId = `copy-${slideIds.length}`
          slideIds.splice(index + 1, 0, slideId)
          return { slideId }
        }),
        snapshotSlide: vi.fn(async (index) => ({
          slideId: slideIds[index]!,
          fingerprint: slideIds[index]!,
        })),
        listSlideShapes: vi.fn(async (index) => ({
          slideId: slideIds[index]!,
          slideIndex: index,
          shapes: [],
        })),
      })
      const proposals = createStructuredProposalController()
      const skill = createPowerPointSkill({ adapter: fake, proposals })
      if (options.readState !== false) await skill.executeTool(call('get_presentation_state'))
      expect(await skill.executeTool(call('plan_deck', { contract }))).not.toHaveProperty(
        'isError',
        true,
      )
      const write = async (name: string, input: Record<string, unknown>) => {
        expect(await skill.executeTool(call(name, input))).not.toHaveProperty('isError', true)
        await proposals.confirm(proposals.pending()!.id)
      }
      const review = async (index: number) => {
        await skill.executeTool(call('screenshot_slide', { slide_index: index }))
        expect(
          await skill.executeTool(
            call('review_slide_screenshot', {
              slide_index: index,
              acceptance_ids: [`A${index + 1}.1`],
              passed: true,
            }),
          ),
        ).not.toHaveProperty('isError', true)
      }
      return { skill, proposals, fake, slideIds, contract, write, review }
    }

    it('appends scaffolds without counting them as produced and requires later fill and review', async () => {
      const { skill, proposals, slideIds, write, review } = await setup()
      await write('edit_slide_text', { slide_index: 0, shape_id: '2', text: 'Hello' })
      const scaffold = await skill.executeTool(call('duplicate_slide', { slide_index: 0 }))
      expect(scaffold.isError).not.toBe(true)
      expect(JSON.parse(scaffold.output).preview).toMatchObject({ scaffold: true })
      await proposals.confirm(proposals.pending()!.id)
      expect(skill.reviewFinalResponse?.({ text: 'Done', mutated: true })).toContain('(1, 2)')
      await expect(
        skill.executeTool(
          call('edit_slide_text', { slide_index: 1, shape_id: '2', text: 'Hello' }),
        ),
      ).resolves.toMatchObject({ isError: true, output: 'design_contract_prototype_required' })
      await skill.executeTool(call('screenshot_slide', { slide_index: 1 }))
      await expect(
        skill.executeTool(
          call('review_slide_screenshot', {
            slide_index: 1,
            acceptance_ids: ['A2.1'],
            passed: true,
          }),
        ),
      ).resolves.toMatchObject({ isError: true, output: 'design_contract_review_not_pending' })

      await write('duplicate_slide', { slide_index: 1 })
      await write('edit_slide_text', { slide_index: 2, shape_id: '2', text: 'Hello' })
      await write('duplicate_slide', { slide_index: 2 })
      await write('edit_slide_text', { slide_index: 3, shape_id: '2', text: 'Hello' })
      expect(slideIds).toHaveLength(4)
      await expect(
        skill.executeTool(
          call('edit_slide_text', { slide_index: 1, shape_id: '2', text: 'Hello' }),
        ),
      ).resolves.toMatchObject({
        isError: true,
        output: expect.stringContaining('"error":"design_contract_review_required"'),
      })
      for (const index of [0, 2, 3]) await review(index)
      await expect(skill.executeTool(call('verify_slides'))).resolves.toMatchObject({
        isError: true,
        output: expect.stringContaining('"error":"design_contract_production_incomplete"'),
      })
      await write('edit_slide_text', { slide_index: 1, shape_id: '2', text: 'Hello' })
      await review(1)
      const verification = await skill.executeTool(call('verify_slides'))
      expect(verification.isError).not.toBe(true)
      expect(JSON.parse(verification.output).status).toBe('verified')
    })

    it.each([
      { label: 'draft contracts', status: 'draft', source: 0 },
      { label: 'non-append duplicates', slideCount: 2, source: 0 },
      { label: 'unknown slide counts', readState: false, source: 0 },
      {
        label: 'pages beyond the last prototype',
        pageCount: 5,
        prototypePages: [1, 2, 3],
        slideCount: 4,
        source: 3,
      },
    ])('does not bypass production gates for $label', async ({ source, ...options }) => {
      const { skill, proposals, fake } = await setup(options)
      await expect(
        skill.executeTool(call('duplicate_slide', { slide_index: source })),
      ).resolves.toMatchObject({ isError: true, output: 'design_contract_prototype_required' })
      expect(proposals.pending()).toBeUndefined()
      expect(fake.duplicateSlide).not.toHaveBeenCalled()
    })

    it('does not record a rejected scaffold as a document mutation', async () => {
      const { skill, proposals, fake } = await setup()
      expect(
        await skill.executeTool(call('duplicate_slide', { slide_index: 0 })),
      ).not.toHaveProperty('isError', true)
      proposals.reject()
      expect(fake.duplicateSlide).not.toHaveBeenCalled()
      expect(skill.buildContext?.()).toContain('"status":"ready"')
      expect(skill.reviewFinalResponse?.({ text: 'Done', mutated: true })).toBeUndefined()
    })

    it('rechecks that scaffold duplication is still an append at confirmation', async () => {
      const { skill, proposals, fake, slideIds } = await setup()
      expect(
        await skill.executeTool(call('duplicate_slide', { slide_index: 0 })),
      ).not.toHaveProperty('isError', true)
      slideIds.push('user-added-slide')
      await expect(proposals.confirm(proposals.pending()!.id)).rejects.toThrow('proposal_stale')
      expect(fake.duplicateSlide).not.toHaveBeenCalled()
    })

    it('invalidates a scaffold proposal when its design contract changes', async () => {
      const { skill, proposals, fake, contract } = await setup()
      expect(
        await skill.executeTool(call('duplicate_slide', { slide_index: 0 })),
      ).not.toHaveProperty('isError', true)
      await skill.executeTool(call('plan_deck', { contract: { ...contract, status: 'draft' } }))
      await expect(proposals.confirm(proposals.pending()!.id)).rejects.toThrow('proposal_stale')
      expect(fake.duplicateSlide).not.toHaveBeenCalled()
    })

    it('does not mark an unverified scaffold as produced', async () => {
      const { skill, proposals, fake } = await setup()
      expect(
        await skill.executeTool(call('duplicate_slide', { slide_index: 0 })),
      ).not.toHaveProperty('isError', true)
      vi.mocked(fake.listSlideShapes).mockResolvedValue({
        slideId: 'wrong-receipt',
        slideIndex: 1,
        shapes: [],
      })
      await expect(proposals.confirm(proposals.pending()!.id)).rejects.toThrow(
        'office_verify_failed',
      )
      expect(skill.buildContext?.()).toContain('"status":"ready"')
      await expect(
        skill.executeTool(call('duplicate_slide', { slide_index: 1 })),
      ).resolves.toMatchObject({ isError: true, output: 'invalid_tool_input' })
      await expect(skill.executeTool(call('verify_slides'))).resolves.toMatchObject({
        isError: true,
        output: expect.stringContaining('"error":"design_contract_production_incomplete"'),
      })
    })

    it.each(['program', 'code'])(
      'directs modern-contract declarative duplicates to the dedicated tool (%s)',
      async (field) => {
        const { skill, proposals, fake, write } = await setup()
        await write('edit_slide_text', { slide_index: 0, shape_id: '2', text: 'Hello' })
        const program = { version: 1, operations: [{ op: 'duplicate_slide', slide_index: 0 }] }
        const result = await skill.executeTool(
          call('execute_office_js', {
            [field]: field === 'program' ? program : JSON.stringify(program),
          }),
        )
        expect(result.isError).toBe(true)
        expect(JSON.parse(result.output)).toMatchObject({
          error: 'invalid_tool_input',
          instruction: expect.stringContaining('duplicate_slide'),
        })
        expect(proposals.pending()).toBeUndefined()
        expect(fake.executeDeclarative).not.toHaveBeenCalled()
      },
    )
  })

  it('requires the exact declarative duplicate receipt instead of accepting any following slide', async () => {
    const fake = adapter({
      executeDeclarative: vi.fn().mockResolvedValue({
        createdShapeIds: [],
        insertedSlideId: 'copy',
      } as never),
      listSlideShapes: vi.fn().mockResolvedValue({
        slideId: 'unrelated-existing-slide',
        slideIndex: 1,
        shapes: [],
      }),
    })
    const proposals = createStructuredProposalController()
    const skill = createPowerPointSkill({ adapter: fake, proposals })
    await skill.executeTool(call('get_presentation_state'))

    await skill.executeTool(
      call('execute_office_js', {
        code: JSON.stringify({
          version: 1,
          operations: [{ op: 'duplicate_slide', slide_index: 0 }],
        }),
      }),
    )

    await expect(proposals.confirm(proposals.pending()!.id)).rejects.toThrow('office_verify_failed')
    await expect(
      skill.executeTool(call('duplicate_slide', { slide_index: 1 })),
    ).resolves.toMatchObject({ isError: true, output: 'invalid_tool_input' })
  })

  it('advertises slide duplication only through the dedicated tool', () => {
    const skill = createPowerPointSkill({
      adapter: adapter(),
      proposals: createStructuredProposalController(),
    })
    const execute = skill.tools.find((tool) => tool.name === 'execute_office_js')!

    expect(JSON.stringify(execute.inputSchema)).not.toContain('duplicate_slide')
    expect(execute.description).toContain('dedicated duplicate_slide tool')
    expect(skill.tools.some((tool) => tool.name === 'duplicate_slide')).toBe(true)
  })

  it('reports the operations field when a declarative duplicate is mixed with edits', async () => {
    const fake = adapter({ executeDeclarative: vi.fn() })
    const skill = createPowerPointSkill({
      adapter: fake,
      proposals: createStructuredProposalController(),
    })

    const result = await skill.executeTool(
      call('execute_office_js', {
        program: {
          version: 1,
          operations: [
            { op: 'duplicate_slide', slide_index: 0 },
            {
              op: 'add_text_box',
              slide_index: 1,
              name: 'title',
              text: 'Title',
              left: 40,
              top: 40,
              width: 600,
              height: 80,
            },
          ],
        },
      }),
    )

    expect(result).toMatchObject({
      isError: true,
      output: 'invalid_tool_input',
      diagnosticError: {
        code: 'InvalidToolInput',
        debugInfo: { errorLocation: 'program.operations' },
      },
    })
    expect(fake.executeDeclarative).not.toHaveBeenCalled()
  })

  it('rejects a declarative mutation that edits and then deletes the same shape', async () => {
    const fake = adapter({ executeDeclarative: vi.fn() })
    const proposals = createStructuredProposalController()
    const skill = createPowerPointSkill({ adapter: fake, proposals })

    await expect(
      skill.executeTool(
        call('execute_office_js', {
          code: JSON.stringify({
            version: 1,
            operations: [
              { op: 'set_shape_text', slide_index: 0, shape_id: '2', text: 'temporary' },
              { op: 'delete_shape', slide_index: 0, shape_id: '2' },
            ],
          }),
        }),
      ),
    ).resolves.toMatchObject({ output: 'invalid_tool_input', isError: true })
    expect(proposals.pending()).toBeUndefined()
    expect(fake.executeDeclarative).not.toHaveBeenCalled()
  })

  it('maps adapter internals to stable errors and validates screenshots', async () => {
    const skill = createPowerPointSkill({
      adapter: adapter({
        listSlideShapes: vi.fn().mockRejectedValue(new Error('secret')),
        screenshotSlide: vi.fn().mockResolvedValue({ mime: 'image/png', base64: 'bad!' }),
      }),
      proposals: createStructuredProposalController(),
    })
    await expect(
      skill.executeTool(call('list_slide_shapes', { slide_index: 0 })),
    ).resolves.toMatchObject({ output: 'office_read_failed', isError: true })
    await expect(
      skill.executeTool(call('screenshot_slide', { slide_index: 0 })),
    ).resolves.toMatchObject({
      output: expect.stringContaining('"error":"office_read_failed"'),
      isError: true,
    })
  })
})

describe('browser PowerPoint adapter', () => {
  it.each([
    [960, 540],
    [720, 405],
  ])('returns the real %i × %i point canvas before production', async (slideWidth, slideHeight) => {
    const slides = { getCount: () => ({ value: 1 }) }
    const pageSetup = { slideWidth, slideHeight, load: vi.fn() }
    vi.stubGlobal('Office', {
      context: { host: 'PowerPoint', requirements: { isSetSupported: () => true } },
    })
    vi.stubGlobal('PowerPoint', {
      run: (fn: (context: unknown) => unknown) =>
        fn({ presentation: { slides, pageSetup }, sync: vi.fn() }),
    })
    expect(await new BrowserPowerPointAdapter().getPresentationState()).toMatchObject({
      slideWidth,
      slideHeight,
      coordinateUnit: 'pt',
    })
    expect(pageSetup.load).toHaveBeenCalledWith(['slideWidth', 'slideHeight'])
    vi.unstubAllGlobals()
  })
  it('reads all actual text instead of an implicit paragraph-marker style', async () => {
    const mixedFont = { load: vi.fn(), name: '', size: 0, color: '', bold: false, italic: false }
    const sampledFont = {
      load: vi.fn(),
      name: 'Aptos',
      size: 24,
      color: '#FFFFFF',
      bold: true,
      italic: false,
    }
    const range = {
      text: 'Title',
      font: mixedFont,
      load: vi.fn(),
      getSubstring: vi.fn(() => ({ font: sampledFont })),
    }
    const shape = { textFrame: { textRange: range } }
    const slide = {
      id: 's1',
      load: vi.fn(),
      shapes: { getItem: vi.fn(() => shape) },
    }
    const context = {
      presentation: {
        slides: { getCount: vi.fn(() => ({ value: 1 })), getItemAt: vi.fn(() => slide) },
      },
      sync: vi.fn().mockResolvedValue(undefined),
    }
    Object.assign(globalThis, {
      Office: {
        context: {
          host: 'PowerPoint',
          requirements: { isSetSupported: vi.fn().mockReturnValue(true) },
        },
      },
      PowerPoint: { run: (callback: (value: typeof context) => unknown) => callback(context) },
    })

    await expect(new BrowserPowerPointAdapter().readShapeTextStyle(0, '2')).resolves.toEqual({
      color: '#FFFFFF',
      fontFamily: 'Aptos',
      fontSize: 24,
      bold: true,
      italic: false,
    })
    expect(range.getSubstring).toHaveBeenCalledWith(0, 5)
    expect(sampledFont.load).toHaveBeenCalledWith('color,name,size,bold,italic')
  })

  it('reads slide count, selected zero-based indices, and the API ladder', async () => {
    const slides = {
      items: [{ id: 'slide-1' }, { id: 'slide-2' }],
      load: vi.fn(),
      getCount: vi.fn(() => ({ value: 2 })),
    }
    const selected = { items: [{ id: 'slide-2' }], load: vi.fn() }
    const isSetSupported = vi.fn((_name: string, version: string) => version !== '1.10')
    Object.assign(globalThis, {
      Office: { context: { host: 'PowerPoint', requirements: { isSetSupported } } },
      PowerPoint: {
        run: (callback: (context: unknown) => unknown) =>
          callback({ presentation: { slides, getSelectedSlides: () => selected }, sync: vi.fn() }),
      },
    })
    await expect(new BrowserPowerPointAdapter().getPresentationState()).resolves.toEqual({
      slideCount: 2,
      selectedSlideIndexes: [1],
      api: { v12: true, v14: true, v15: true, v18: true, v110: false },
    })
    expect(slides.load).toHaveBeenCalledWith('items/id')
  })

  it('keeps presentation state usable when the optional 1.5 selection read fails', async () => {
    const slides = {
      items: [{ id: 'slide-1' }, { id: 'slide-2' }],
      load: vi.fn(),
      getCount: vi.fn(() => ({ value: 2 })),
    }
    const selected = { items: [], load: vi.fn() }
    const sync = vi
      .fn()
      .mockResolvedValueOnce(undefined)
      .mockRejectedValueOnce(new Error('host busy'))
    Object.assign(globalThis, {
      Office: {
        context: { host: 'PowerPoint', requirements: { isSetSupported: vi.fn(() => true) } },
      },
      PowerPoint: {
        run: (callback: (context: unknown) => unknown) =>
          callback({
            presentation: { slides, getSelectedSlides: () => selected },
            sync,
          }),
      },
    })

    await expect(new BrowserPowerPointAdapter().getPresentationState()).resolves.toEqual({
      slideCount: 2,
      selectedSlideIndexes: [],
      api: { v12: true, v14: true, v15: true, v18: true, v110: true },
    })
    expect(sync).toHaveBeenCalledTimes(2)
  })

  it('falls back to a bounded slide collection when Mac rejects getCount sync', async () => {
    const slides = {
      items: [{ id: 'slide-1' }, { id: 'slide-2' }],
      load: vi.fn(),
      getCount: vi.fn(() => ({ value: 2 })),
    }
    const sync = vi
      .fn()
      .mockRejectedValueOnce(new Error('GeneralException'))
      .mockResolvedValue(undefined)
    Object.assign(globalThis, {
      Office: {
        context: {
          host: 'PowerPoint',
          platform: 'Mac',
          requirements: { isSetSupported: vi.fn(() => true) },
        },
      },
      PowerPoint: {
        run: (callback: (context: unknown) => unknown) =>
          callback({ presentation: { slides }, sync }),
      },
    })

    await expect(new BrowserPowerPointAdapter().getPresentationState()).resolves.toMatchObject({
      slideCount: 2,
    })
    expect(slides.load).toHaveBeenCalledWith('items/id')
    expect(sync).toHaveBeenCalledTimes(2)
  })

  it('reads slide count from the compressed document when Mac rejects both JS API paths', async () => {
    const archive = new JSZip()
    archive.file('ppt/slides/slide1.xml', '<p:sld/>')
    archive.file('ppt/slides/slide2.xml', '<p:sld/>')
    archive.file('ppt/slideLayouts/slideLayout1.xml', '<p:sldLayout/>')
    const bytes = await archive.generateAsync({ type: 'uint8array' })
    const closeAsync = vi.fn((callback: () => void) => callback())
    const getFileAsync = vi.fn(
      (_type: unknown, _options: unknown, callback: (result: unknown) => void) =>
        callback({
          status: 'succeeded',
          value: {
            size: bytes.byteLength,
            sliceCount: 1,
            getSliceAsync: (_index: number, done: (result: unknown) => void) =>
              done({ status: 'succeeded', value: { data: Array.from(bytes) } }),
            closeAsync,
          },
        }),
    )
    const slides = { items: [], load: vi.fn(), getCount: vi.fn(() => ({ value: 2 })) }
    Object.assign(globalThis, {
      Office: {
        FileType: { Compressed: 'compressed' },
        context: {
          host: 'PowerPoint',
          platform: 'Mac',
          document: { getFileAsync },
          requirements: { isSetSupported: vi.fn(() => true) },
        },
      },
      PowerPoint: {
        run: (callback: (context: unknown) => unknown) =>
          callback({
            presentation: { slides },
            sync: vi.fn().mockRejectedValue(new Error('GeneralException')),
          }),
      },
    })

    await expect(new BrowserPowerPointAdapter().getPresentationState()).resolves.toMatchObject({
      slideCount: 2,
    })
    expect(getFileAsync).toHaveBeenCalledWith(
      'compressed',
      { sliceSize: 4 * 1024 * 1024 },
      expect.any(Function),
    )
    expect(closeAsync).toHaveBeenCalledOnce()
  })

  it('reads bounded state on PowerPointApi 1.2 hosts through getCount only', async () => {
    const count = { value: 1 }
    const slides = { getCount: vi.fn(() => count) }
    const isSetSupported = vi.fn(
      (_name: string, version: string) => version === '1.2' || version === '1.4',
    )
    Object.assign(globalThis, {
      Office: { context: { host: 'PowerPoint', requirements: { isSetSupported } } },
      PowerPoint: {
        run: (callback: (context: unknown) => unknown) =>
          callback({ presentation: { slides }, sync: vi.fn() }),
      },
    })

    await expect(new BrowserPowerPointAdapter().getPresentationState()).resolves.toEqual({
      slideCount: 1,
      selectedSlideIndexes: [],
      api: { v12: true, v14: true, v15: false, v18: false, v110: false },
    })
    expect(slides.getCount).toHaveBeenCalledOnce()
  })

  it('rejects master package replacement on Mac before entering PowerPoint.run', async () => {
    const run = vi.fn()
    Object.assign(globalThis, {
      Office: {
        context: {
          host: 'PowerPoint',
          platform: 'Mac',
          requirements: { isSetSupported: vi.fn().mockReturnValue(true) },
        },
      },
      PowerPoint: { run },
    })

    await expect(
      new BrowserPowerPointAdapter().replaceSlidePackage(0, 'ppt', true),
    ).rejects.toThrow('office_api_unsupported')
    expect(run).not.toHaveBeenCalled()
  })

  const originals = { Office: globalThis.Office, PowerPoint: globalThis.PowerPoint }
  afterEach(() => Object.assign(globalThis, originals))

  it('detects host/API support before PowerPoint.run', async () => {
    const run = vi.fn()
    Object.assign(globalThis, {
      Office: {
        context: { host: 'Word', requirements: { isSetSupported: vi.fn().mockReturnValue(true) } },
      },
      PowerPoint: { run },
    })
    await expect(new BrowserPowerPointAdapter().listSlideShapes(0)).rejects.toThrow(
      'office_api_unsupported',
    )
    expect(run).not.toHaveBeenCalled()
  })

  it('requires the per-operation PowerPoint API set', async () => {
    const run = vi.fn()
    const supports = vi.fn((_name: string, version: string) => version === '1.4')
    Object.assign(globalThis, {
      Office: { context: { host: 'PowerPoint', requirements: { isSetSupported: supports } } },
      PowerPoint: { run },
    })
    await expect(new BrowserPowerPointAdapter().screenshotSlide(0)).rejects.toThrow(
      'office_api_unsupported',
    )
    await expect(new BrowserPowerPointAdapter().verifySlides()).rejects.toThrow(
      'office_api_unsupported',
    )
    expect(run).not.toHaveBeenCalled()
    expect(supports).toHaveBeenCalledWith('PowerPointApi', '1.8')
    expect(supports).toHaveBeenCalledWith('PowerPointApi', '1.10')
  })

  it('retries a transient native screenshot failure inside one tool call', async () => {
    const image = { value: png }
    const slide = { id: 's1', load: vi.fn(), getImageAsBase64: vi.fn(() => image) }
    const slides = {
      getCount: vi.fn(() => ({ value: 1 })),
      getItemAt: vi.fn(() => slide),
    }
    const context = { presentation: { slides }, sync: vi.fn().mockResolvedValue(undefined) }
    const run = vi
      .fn()
      .mockRejectedValueOnce(new Error('GeneralException'))
      .mockImplementation((callback: (value: typeof context) => unknown) => callback(context))
    Object.assign(globalThis, {
      Office: {
        context: {
          host: 'PowerPoint',
          requirements: { isSetSupported: vi.fn().mockReturnValue(true) },
        },
      },
      PowerPoint: { run },
    })

    await expect(new BrowserPowerPointAdapter().screenshotSlide(0)).resolves.toEqual({
      base64: png,
      mime: 'image/png',
    })
    expect(run).toHaveBeenCalledTimes(2)
  })

  it('switches between documented rendering shapes when Mac rejects sized screenshots', async () => {
    const getImageAsBase64 = vi.fn((options?: { width?: number; height?: number }) => {
      if (options !== undefined) throw new Error('GeneralException')
      return { value: png }
    })
    const slide = { id: 's1', load: vi.fn(), getImageAsBase64 }
    const slides = {
      getCount: vi.fn(() => ({ value: 1 })),
      getItemAt: vi.fn(() => slide),
    }
    const context = { presentation: { slides }, sync: vi.fn().mockResolvedValue(undefined) }
    Object.assign(globalThis, {
      Office: {
        context: {
          host: 'PowerPoint',
          platform: 'Mac',
          requirements: { isSetSupported: vi.fn().mockReturnValue(true) },
        },
      },
      PowerPoint: {
        run: (callback: (value: typeof context) => unknown) => callback(context),
      },
    })

    await expect(new BrowserPowerPointAdapter().screenshotSlide(0)).resolves.toEqual({
      base64: png,
      mime: 'image/png',
    })
    expect(getImageAsBase64).toHaveBeenNthCalledWith(1, { width: 960 })
    expect(getImageAsBase64).toHaveBeenNthCalledWith(2, { height: 540 })
    expect(getImageAsBase64).toHaveBeenNthCalledWith(3, undefined)
  })

  it('retries empty or invalid native screenshot data before giving up rendering', async () => {
    const getImageAsBase64 = vi
      .fn()
      .mockReturnValueOnce({ value: '' })
      .mockReturnValueOnce({ value: 'iVBORw0KGgoAAAA=' })
      .mockReturnValueOnce({ value: png })
    const slide = { id: 's1', load: vi.fn(), getImageAsBase64 }
    const context = {
      presentation: {
        slides: { getCount: () => ({ value: 1 }), getItemAt: () => slide },
      },
      sync: vi.fn().mockResolvedValue(undefined),
    }
    vi.stubGlobal('Office', {
      context: { host: 'PowerPoint', requirements: { isSetSupported: () => true } },
    })
    vi.stubGlobal('PowerPoint', {
      run: (callback: (value: typeof context) => unknown) => callback(context),
    })
    await expect(new BrowserPowerPointAdapter().screenshotSlide(0)).resolves.toEqual({
      mime: 'image/png',
      base64: png,
    })
    expect(getImageAsBase64).toHaveBeenCalledTimes(3)
  })

  describe('native screenshot lifetime', () => {
    const stages = [
      'run admission',
      'slide count sync',
      'slide lookup sync',
      'image export sync',
      'run cleanup',
    ] as const

    function stallScreenshot(stage: (typeof stages)[number]) {
      let release!: () => void
      const blocked = new Promise<void>((resolve) => {
        release = resolve
      })
      const getImageAsBase64 = vi.fn(() => ({ value: png }))
      const slide = { id: 'slide-1', load: vi.fn(), getImageAsBase64 }
      let syncCount = 0
      const context = {
        presentation: {
          slides: { getCount: () => ({ value: 1 }), getItemAt: () => slide },
        },
        sync: vi.fn(async () => {
          syncCount += 1
          if (
            (stage === 'slide count sync' && syncCount === 1) ||
            (stage === 'slide lookup sync' && syncCount === 2) ||
            (stage === 'image export sync' && syncCount === 3)
          )
            await blocked
        }),
      }
      const run = vi.fn(async (callback: (value: typeof context) => Promise<unknown>) => {
        if (stage === 'run admission') await blocked
        const result = await callback(context)
        if (stage === 'run cleanup') await blocked
        return result
      })
      vi.stubGlobal('Office', {
        context: { host: 'PowerPoint', requirements: { isSetSupported: () => true } },
      })
      vi.stubGlobal('PowerPoint', { run })
      return { release, run, getImageAsBase64 }
    }

    afterEach(() => {
      vi.unstubAllGlobals()
      vi.useRealTimers()
    })

    it.each(stages)('bounds a stalled %s within the native read budget', async (stage) => {
      vi.useFakeTimers()
      const native = stallScreenshot(stage)
      let outcome: string | undefined
      const pending = new BrowserPowerPointAdapter().screenshotSlide(0).then(
        () => {
          outcome = 'success'
        },
        (error: Error) => {
          outcome = error.message
        },
      )
      try {
        // Native capture must leave room for the existing 10s preview inside the PC tool limit.
        await vi.advanceTimersByTimeAsync(15_000)
        expect(outcome).toBe('office_screenshot_unavailable')
        expect(native.run).toHaveBeenCalledOnce()
        expect(vi.getTimerCount()).toBe(0)
        const exportsBeforeRelease = native.getImageAsBase64.mock.calls.length
        native.release()
        await pending
        await vi.advanceTimersByTimeAsync(0)
        expect(native.getImageAsBase64).toHaveBeenCalledTimes(exportsBeforeRelease)
        expect(native.run).toHaveBeenCalledOnce()
      } finally {
        native.release()
        await pending
      }
    })

    it.each(stages)('settles user cancellation while %s is still stalled', async (stage) => {
      vi.useFakeTimers()
      const native = stallScreenshot(stage)
      const controller = new AbortController()
      let outcome: string | undefined
      const pending = new BrowserPowerPointAdapter().screenshotSlide(0, controller.signal).then(
        () => {
          outcome = 'success'
        },
        (error: Error) => {
          outcome = error.message
        },
      )
      try {
        await vi.advanceTimersByTimeAsync(0)
        controller.abort()
        await vi.advanceTimersByTimeAsync(0)
        expect(outcome).toBe('cancelled')
        expect(native.run).toHaveBeenCalledOnce()
        expect(vi.getTimerCount()).toBe(0)
      } finally {
        native.release()
        await pending
      }
    })

    it('shares one native read deadline across concrete rendering failures', async () => {
      vi.useFakeTimers()
      const native = stallScreenshot('image export sync')
      native.run.mockImplementationOnce(async () => {
        await new Promise((resolve) => setTimeout(resolve, 10_000))
        throw new Error('GeneralException')
      })
      let outcome: string | undefined
      const pending = new BrowserPowerPointAdapter().screenshotSlide(0).then(
        () => {
          outcome = 'success'
        },
        (error: Error) => {
          outcome = error.message
        },
      )
      try {
        await vi.advanceTimersByTimeAsync(14_999)
        expect(native.run).toHaveBeenCalledTimes(2)
        expect(outcome).toBeUndefined()
        await vi.advanceTimersByTimeAsync(1)
        expect(outcome).toBe('office_screenshot_unavailable')
        expect(native.run).toHaveBeenCalledTimes(2)
      } finally {
        native.release()
        await pending
      }
    })

    it('rejects an overdue native result even when the deadline timer has not run yet', async () => {
      vi.useFakeTimers()
      const native = stallScreenshot('run cleanup')
      const pending = new BrowserPowerPointAdapter().screenshotSlide(0).then(
        () => 'success',
        (error: Error) => error.message,
      )
      try {
        await vi.advanceTimersByTimeAsync(0)
        vi.setSystemTime(Date.now() + 15_001)
        native.release()
        expect(await pending).toBe('office_screenshot_unavailable')
      } finally {
        native.release()
        await pending
      }
    })

    it('does not let a late timed-out native screenshot unlock visual review', async () => {
      vi.useFakeTimers()
      const native = stallScreenshot('image export sync')
      const browser = new BrowserPowerPointAdapter()
      const proposals = createStructuredProposalController()
      const skill = createPowerPointSkill({
        adapter: adapter({ screenshotSlide: browser.screenshotSlide.bind(browser) }),
        proposals,
      })
      await skill.executeTool(call('plan_deck', { contract: modernContract() }))
      await skill.executeTool(
        call('edit_slide_text', { slide_index: 0, shape_id: '2', text: 'Hello' }),
      )
      await proposals.confirm(proposals.pending()!.id)
      const pending = skill.executeTool(call('screenshot_slide', { slide_index: 0 }))
      try {
        await vi.advanceTimersByTimeAsync(15_000)
        native.release()
        const screenshot = await pending
        expect.soft(screenshot).toMatchObject({ isError: true })
        await expect(
          skill.executeTool(
            call('review_slide_screenshot', {
              slide_index: 0,
              acceptance_ids: ['A1.1'],
              passed: true,
            }),
          ),
        ).resolves.toMatchObject({ isError: true, output: 'design_contract_screenshot_required' })
      } finally {
        native.release()
        await pending
      }
    })
  })

  it('maps native master operations to PowerPointApi 1.10 objects', async () => {
    const setSolidFill = vi.fn()
    const setThemeColor = vi.fn()
    const layoutBackground: Record<string, unknown> = {}
    const layout = { background: layoutBackground }
    const master = {
      background: { fill: { setSolidFill } },
      themeColorScheme: { setThemeColor },
      layouts: { getItem: vi.fn().mockReturnValue(layout) },
    }
    const context = {
      presentation: { slideMasters: { getItem: vi.fn().mockReturnValue(master) } },
      sync: vi.fn().mockResolvedValue(undefined),
    }
    Object.assign(globalThis, {
      Office: {
        context: {
          host: 'PowerPoint',
          requirements: { isSetSupported: vi.fn().mockReturnValue(true) },
        },
      },
      PowerPoint: { run: (callback: (value: typeof context) => unknown) => callback(context) },
    })

    await new BrowserPowerPointAdapter().executeMasterOperations([
      {
        op: 'set_master_background',
        master_id: 'm1',
        fill: { type: 'solid', color: '#000000', transparency: 0.2 },
      },
      {
        op: 'set_master_theme_color',
        master_id: 'm1',
        theme_color: 'Light1',
        color: '#FFFFFF',
      },
      {
        op: 'set_layout_background_following',
        master_id: 'm1',
        layout_id: 'l1',
        follow_master: true,
        show_master_graphics: false,
      },
    ])

    expect(setSolidFill).toHaveBeenCalledWith({ color: '#000000', transparency: 0.2 })
    expect(setThemeColor).toHaveBeenCalledWith('Light1', '#FFFFFF')
    expect(layoutBackground).toMatchObject({
      isMasterBackgroundFollowed: true,
      areBackgroundGraphicsHidden: true,
    })
  })

  it('sets and verifies a native per-slide solid background', async () => {
    const fake = adapter({
      readSlideBackground: vi
        .fn()
        .mockResolvedValueOnce({
          slideId: 'slide-1',
          type: 'Solid',
          backgroundColor: '#FFFFFF',
          transparency: 0,
        })
        .mockResolvedValueOnce({
          slideId: 'slide-1',
          type: 'Solid',
          backgroundColor: '#FFFFFF',
          transparency: 0,
        })
        .mockResolvedValue({
          slideId: 'slide-1',
          type: 'Solid',
          backgroundColor: '#112233',
          transparency: 0,
        }),
      setSlideBackground: vi.fn().mockResolvedValue(undefined),
    })
    const proposals = createStructuredProposalController()
    const skill = createPowerPointSkill({ adapter: fake, proposals })

    await expect(
      skill.executeTool(
        call('set_slide_background', {
          slide_index: 0,
          color: '#112233',
          transparency: 0,
        }),
      ),
    ).resolves.toMatchObject({ output: expect.stringContaining('set_slide_background') })
    await expect(proposals.confirm(proposals.pending()!.id)).resolves.toBeUndefined()
    expect(fake.setSlideBackground).toHaveBeenCalledWith(0, '#112233', 0, expect.any(AbortSignal))
  })

  it('maps per-slide backgrounds to the PowerPointApi 1.10 fill object', async () => {
    const setSolidFill = vi.fn()
    const solid = {
      isNullObject: false,
      color: '#ABCDEF',
      transparency: 0.25,
      load: vi.fn(),
    }
    const fill = {
      type: 'Solid',
      load: vi.fn(),
      getSolidFillOrNullObject: vi.fn().mockReturnValue(solid),
      setSolidFill,
    }
    const slide = { id: 'slide-1', load: vi.fn(), background: { fill } }
    const context = {
      presentation: {
        slides: {
          getCount: vi.fn().mockReturnValue({ value: 1 }),
          getItemAt: vi.fn().mockReturnValue(slide),
        },
      },
      sync: vi.fn().mockResolvedValue(undefined),
    }
    Object.assign(globalThis, {
      Office: {
        context: {
          host: 'PowerPoint',
          requirements: { isSetSupported: vi.fn().mockReturnValue(true) },
        },
      },
      PowerPoint: { run: (callback: (value: typeof context) => unknown) => callback(context) },
    })
    const subject = new BrowserPowerPointAdapter()

    await expect(subject.readSlideBackground(0)).resolves.toEqual({
      slideId: 'slide-1',
      type: 'Solid',
      backgroundColor: '#ABCDEF',
      transparency: 0.25,
    })
    await subject.setSlideBackground(0, '#112233', 0)
    expect(setSolidFill).toHaveBeenCalledWith({ color: '#112233', transparency: 0 })
  })

  it('routes declarative duplication through the transaction-safe duplicate primitive', async () => {
    const subject = new BrowserPowerPointAdapter()
    const duplicate = vi.spyOn(subject, 'duplicateSlide').mockResolvedValue({ slideId: 'copy' })

    await expect(
      subject.executeDeclarative([{ op: 'duplicate_slide', slide_index: 0 }]),
    ).resolves.toEqual({ createdShapeIds: [], insertedSlideId: 'copy' })
    expect(duplicate).toHaveBeenCalledWith(0, undefined)
  })

  it('defensively rejects edit-then-delete before entering PowerPoint.run', async () => {
    const run = vi.fn()
    Object.assign(globalThis, {
      Office: {
        context: {
          host: 'PowerPoint',
          requirements: { isSetSupported: vi.fn().mockReturnValue(true) },
        },
      },
      PowerPoint: { run },
    })

    await expect(
      new BrowserPowerPointAdapter().executeDeclarative([
        { op: 'set_shape_text', slide_index: 0, shape_id: '2', text: 'temporary' },
        { op: 'delete_shape', slide_index: 0, shape_id: '2' },
      ]),
    ).rejects.toThrow('invalid_tool_input')
    expect(run).not.toHaveBeenCalled()
  })

  it('returns stable IDs/geometry and verifies negative, overflow, and overlap geometry', async () => {
    const sync = vi.fn().mockResolvedValue(undefined)
    const shapes = {
      load: vi.fn(),
      items: [
        { id: '2', name: 'A', type: 'TextBox', left: -5, top: 10, width: 100, height: 50 },
        { id: '3', name: 'B', type: 'TextBox', left: 50, top: 20, width: 950, height: 530 },
      ],
    }
    const slides = {
      load: vi.fn(),
      items: [{ id: 's1', shapes, load: vi.fn() }],
      getCount: vi.fn(() => ({ value: 1 })),
      getItemAt: vi.fn((i) => slides.items[i]),
    }
    Object.assign(globalThis, {
      Office: {
        context: {
          host: 'PowerPoint',
          requirements: { isSetSupported: vi.fn().mockReturnValue(true) },
        },
      },
      PowerPoint: {
        run: (callback: (context: unknown) => unknown) =>
          callback({
            presentation: {
              slides,
              pageSetup: { slideWidth: 960, slideHeight: 540, load: vi.fn() },
            },
            sync,
          }),
      },
    })
    await expect(new BrowserPowerPointAdapter().listSlideShapes(0)).resolves.toMatchObject({
      slideId: 's1',
      shapes: [
        { id: '2', left: -5 },
        { id: '3', left: 50 },
      ],
    })
    await expect(new BrowserPowerPointAdapter().verifySlides()).resolves.toMatchObject({
      slides: [
        {
          overflows: expect.arrayContaining([
            expect.objectContaining({ shapeId: '2', edge: 'left' }),
            expect.objectContaining({ shapeId: '3', edge: 'right' }),
            expect.objectContaining({ shapeId: '3', edge: 'bottom' }),
          ]),
          overlaps: [{ shapeAId: '2', shapeBId: '3', overlapX: 45, overlapY: 40 }],
        },
      ],
    })
    expect(sync).toHaveBeenCalled()
  })

  it('does not report text intentionally placed inside a full-bleed image as an overlap', async () => {
    const sync = vi.fn().mockResolvedValue(undefined)
    const shapes = {
      load: vi.fn(),
      items: [
        { id: '4', name: 'Cover', type: 'Image', left: 0, top: 0, width: 960, height: 540 },
        { id: '5', name: 'Title', type: 'TextBox', left: 80, top: 70, width: 330, height: 24 },
        { id: '6', name: 'Body', type: 'TextBox', left: 80, top: 120, width: 380, height: 132 },
        { id: '7', name: 'Clipped', type: 'TextBox', left: 930, top: 300, width: 60, height: 30 },
      ],
    }
    const slides = { load: vi.fn(), items: [{ id: 's1', shapes, load: vi.fn() }] }
    Object.assign(globalThis, {
      Office: {
        context: {
          host: 'PowerPoint',
          requirements: { isSetSupported: vi.fn().mockReturnValue(true) },
        },
      },
      PowerPoint: {
        run: (callback: (context: unknown) => unknown) =>
          callback({
            presentation: {
              slides,
              pageSetup: { slideWidth: 960, slideHeight: 540, load: vi.fn() },
            },
            sync,
          }),
      },
    })

    await expect(new BrowserPowerPointAdapter().verifySlides()).resolves.toMatchObject({
      slides: [
        {
          overlaps: [{ shapeAId: '4', shapeBId: '7', overlapX: 30, overlapY: 30 }],
          overflows: [expect.objectContaining({ shapeId: '7', edge: 'right' })],
        },
      ],
    })
  })

  it('retries shape inventory with basic fields when Mac rejects rich geometry loading', async () => {
    const sync = vi
      .fn()
      .mockResolvedValueOnce(undefined)
      .mockResolvedValueOnce(undefined)
      .mockRejectedValueOnce(new Error('GeneralException'))
      .mockResolvedValue(undefined)
    const shapes = {
      load: vi.fn(),
      items: [{ id: '2', name: 'Title', type: 'TextBox' }],
    }
    const slide = { id: 's1', shapes, load: vi.fn() }
    const slides = {
      getCount: vi.fn(() => ({ value: 1 })),
      getItemAt: vi.fn(() => slide),
    }
    Object.assign(globalThis, {
      Office: {
        context: {
          host: 'PowerPoint',
          platform: 'Mac',
          requirements: { isSetSupported: vi.fn().mockReturnValue(true) },
        },
      },
      PowerPoint: {
        run: (callback: (context: unknown) => unknown) =>
          callback({ presentation: { slides }, sync }),
      },
    })

    await expect(new BrowserPowerPointAdapter().listSlideShapes(0)).resolves.toMatchObject({
      slideId: 's1',
      shapes: [{ id: '2', name: 'Title', type: 'TextBox' }],
    })
    expect(shapes.load).toHaveBeenLastCalledWith('items/id,items/name,items/type')
  })

  it('checks cancellation before every write/sync and implements text edit and duplicate', async () => {
    const packageZip = new JSZip()
    packageZip.file('ppt/slides/slide1.xml', '<p:sld xmlns:p="urn:p"/>')
    const slidePackage = await packageZip.generateAsync({ type: 'base64' })
    const sync = vi.fn().mockResolvedValue(undefined)
    const textRange = { text: 'Old', load: vi.fn() }
    const shape = {
      id: '2',
      name: 'Title',
      type: 'Placeholder',
      left: 10,
      top: 20,
      width: 200,
      height: 40,
      textFrame: { hasText: true, load: vi.fn(), textRange },
    }
    const slide = {
      id: 's1',
      load: vi.fn(),
      shapes: { load: vi.fn(), items: [shape], getItem: vi.fn(() => shape) },
      exportAsBase64: vi.fn(() => ({ value: slidePackage })),
    }
    const slides = {
      load: vi.fn(),
      items: [slide],
      getCount: vi.fn(() => ({ value: slides.items.length })),
      getItemAt: vi.fn((index: number) => slides.items[index]),
    }
    const insertSlidesFromBase64 = vi.fn(() => {
      slides.items.splice(1, 0, { ...slide, id: 's2' })
    })
    Object.assign(globalThis, {
      Office: {
        context: {
          host: 'PowerPoint',
          requirements: { isSetSupported: vi.fn().mockReturnValue(true) },
        },
      },
      PowerPoint: {
        run: (callback: (context: unknown) => unknown) =>
          callback({ presentation: { slides, insertSlidesFromBase64 }, sync }),
      },
    })
    const subject = new BrowserPowerPointAdapter()
    await expect(subject.snapshotSlide(0)).resolves.toMatchObject({
      slideId: 's1',
      fingerprint: expect.stringMatching(/^s1:\d+:[0-9a-f]{8}$/),
    })
    await subject.editSlideText(0, '2', 'New')
    expect(textRange.text).toBe('New')
    await expect(subject.duplicateSlide(0)).resolves.toEqual({ slideId: 's2' })
    expect(insertSlidesFromBase64).toHaveBeenCalledWith(slidePackage, { targetSlideId: 's1' })

    const controller = new AbortController()
    controller.abort()
    await expect(subject.editSlideText(0, '2', 'No', controller.signal)).rejects.toThrow(
      'cancelled',
    )
    expect(textRange.text).toBe('New')
  })

  it('accepts the vertical-tab paragraph separator returned by Mac PowerPoint after a write', async () => {
    const textRange = { text: 'Old', load: vi.fn() }
    const shape = { textFrame: { textRange } }
    const slide = {
      id: 's1',
      load: vi.fn(),
      shapes: { getItem: vi.fn(() => shape) },
    }
    const slides = {
      getCount: vi.fn(() => ({ value: 1 })),
      getItemAt: vi.fn(() => slide),
    }
    const sync = vi.fn().mockImplementation(async () => {
      if (textRange.text === 'First line\nSecond line') textRange.text = 'First line\vSecond line'
    })
    Object.assign(globalThis, {
      Office: {
        context: {
          host: 'PowerPoint',
          requirements: { isSetSupported: vi.fn().mockReturnValue(true) },
        },
      },
      PowerPoint: {
        run: (callback: (context: unknown) => unknown) =>
          callback({ presentation: { slides }, sync }),
      },
    })

    await expect(
      new BrowserPowerPointAdapter().executeDeclarative([
        {
          op: 'set_shape_text',
          slide_index: 0,
          shape_id: '2',
          text: 'First line\nSecond line',
        },
      ]),
    ).resolves.toEqual({ createdShapeIds: [] })
  })

  it('reconciles a text sync rejection that committed and never overwrites a third state', async () => {
    const textRange = { text: 'Old', load: vi.fn() }
    const shape = { textFrame: { textRange } }
    const slide = {
      id: 's1',
      load: vi.fn(),
      shapes: { getItem: vi.fn(() => shape) },
    }
    const slides = {
      getCount: vi.fn(() => ({ value: 1 })),
      getItemAt: vi.fn(() => slide),
    }
    let rejected = false
    const sync = vi.fn().mockImplementation(async () => {
      if (textRange.text === 'New' && !rejected) {
        rejected = true
        throw new Error('host rejected after commit')
      }
    })
    Object.assign(globalThis, {
      Office: {
        context: {
          host: 'PowerPoint',
          requirements: { isSetSupported: vi.fn().mockReturnValue(true) },
        },
      },
      PowerPoint: {
        run: (callback: (context: unknown) => unknown) =>
          callback({ presentation: { slides }, sync }),
      },
    })

    await expect(
      new BrowserPowerPointAdapter().editSlideText(0, '2', 'New'),
    ).resolves.toBeUndefined()
    expect(textRange.text).toBe('New')

    textRange.text = 'Old'
    sync.mockClear()
    rejected = false
    sync.mockImplementation(async () => {
      if (textRange.text === 'New' && !rejected) {
        rejected = true
        textRange.text = 'User edit'
        throw new Error('host rejected after third-party edit')
      }
    })
    await expect(new BrowserPowerPointAdapter().editSlideText(0, '2', 'New')).rejects.toThrow(
      'office_concurrent_change',
    )
    expect(textRange.text).toBe('User edit')
  })

  it('uses bounded convergence when a rejected text sync is initially read back stale', async () => {
    let visible = 'Old'
    let assigned = 'Old'
    let reconciliationReads = 0
    const textRange = {
      load: vi.fn(),
      get text() {
        return visible
      },
      set text(value: string) {
        assigned = value
      },
    }
    const slide = {
      id: 's1',
      load: vi.fn(),
      shapes: { getItem: vi.fn(() => ({ textFrame: { textRange } })) },
    }
    const slides = {
      getCount: vi.fn(() => ({ value: 1 })),
      getItemAt: vi.fn(() => slide),
    }
    let rejected = false
    const sync = vi.fn().mockImplementation(async () => {
      if (assigned === 'New' && !rejected) {
        rejected = true
        throw new Error('host rejected after commit')
      }
      if (rejected && visible !== assigned && ++reconciliationReads >= 3) visible = assigned
    })
    Object.assign(globalThis, {
      Office: {
        context: {
          host: 'PowerPoint',
          requirements: { isSetSupported: vi.fn().mockReturnValue(true) },
        },
      },
      PowerPoint: {
        run: (callback: (context: unknown) => unknown) =>
          callback({ presentation: { slides }, sync }),
      },
    })

    await expect(
      new BrowserPowerPointAdapter().editSlideText(0, '2', 'New'),
    ).resolves.toBeUndefined()
    expect(visible).toBe('New')
    expect(reconciliationReads).toBe(3)
  })

  it('uses bounded convergence after a rejected declarative sync', async () => {
    let visible = 'Old'
    let assigned = 'Old'
    let reconciliationReads = 0
    const textRange = {
      load: vi.fn(),
      get text() {
        return visible
      },
      set text(value: string) {
        assigned = value
      },
    }
    const shape = { id: '2', load: vi.fn(), textFrame: { textRange } }
    const slide = {
      id: 's1',
      load: vi.fn(),
      shapes: { getItem: vi.fn(() => shape) },
    }
    const slides = {
      getCount: vi.fn(() => ({ value: 1 })),
      getItemAt: vi.fn(() => slide),
    }
    let rejected = false
    const sync = vi.fn().mockImplementation(async () => {
      if (assigned === 'New' && !rejected) {
        rejected = true
        throw new Error('host rejected after declarative commit')
      }
      if (rejected && visible !== assigned && ++reconciliationReads >= 3) visible = assigned
    })
    Object.assign(globalThis, {
      Office: {
        context: {
          host: 'PowerPoint',
          requirements: { isSetSupported: vi.fn().mockReturnValue(true) },
        },
      },
      PowerPoint: {
        run: (callback: (context: unknown) => unknown) =>
          callback({ presentation: { slides }, sync }),
      },
    })

    await expect(
      new BrowserPowerPointAdapter().executeDeclarative([
        { op: 'set_shape_text', slide_index: 0, shape_id: '2', text: 'New' },
      ]),
    ).resolves.toEqual({ createdShapeIds: [] })
    expect(visible).toBe('New')
  })

  it('reports an attributable text commit as applied when cancellation races after sync', async () => {
    const controller = new AbortController()
    const textRange = { text: 'Old', load: vi.fn() }
    const shape = { textFrame: { textRange } }
    const slide = {
      id: 's1',
      load: vi.fn(),
      shapes: { getItem: vi.fn(() => shape) },
    }
    const slides = {
      getCount: vi.fn(() => ({ value: 1 })),
      getItemAt: vi.fn(() => slide),
    }
    let aborted = false
    const sync = vi.fn().mockImplementation(async () => {
      if (textRange.text === 'New' && !aborted) {
        aborted = true
        controller.abort()
      }
    })
    Object.assign(globalThis, {
      Office: {
        context: {
          host: 'PowerPoint',
          requirements: { isSetSupported: vi.fn().mockReturnValue(true) },
        },
      },
      PowerPoint: {
        run: (callback: (context: unknown) => unknown) =>
          callback({ presentation: { slides }, sync }),
      },
    })

    await expect(
      new BrowserPowerPointAdapter().editSlideText(0, '2', 'New', controller.signal),
    ).resolves.toBeUndefined()
    expect(textRange.text).toBe('New')
  })

  it('converges delayed text visibility even when cancellation races readback', async () => {
    const runScenario = async (abortDuringReadback: boolean) => {
      const controller = new AbortController()
      let visible = 'Old'
      let pending: string | undefined
      let readbacks = 0
      const textRange = {
        load: vi.fn(),
        get text() {
          return visible
        },
        set text(value: string) {
          pending = value
        },
      }
      const shape = { textFrame: { textRange } }
      const slide = {
        id: 's1',
        load: vi.fn(),
        shapes: { getItem: vi.fn(() => shape) },
      }
      const slides = {
        getCount: vi.fn(() => ({ value: 1 })),
        getItemAt: vi.fn(() => slide),
      }
      const sync = vi.fn().mockImplementation(async () => {
        if (pending === 'New') {
          readbacks += 1
          if (abortDuringReadback && readbacks === 2) controller.abort()
          if (readbacks >= 3) visible = pending
        } else if (pending === 'Old') visible = pending
      })
      Object.assign(globalThis, {
        Office: {
          context: {
            host: 'PowerPoint',
            requirements: { isSetSupported: vi.fn().mockReturnValue(true) },
          },
        },
        PowerPoint: {
          run: (callback: (context: unknown) => unknown) =>
            callback({ presentation: { slides }, sync }),
        },
      })
      const result = new BrowserPowerPointAdapter().editSlideText(0, '2', 'New', controller.signal)
      if (abortDuringReadback) {
        await expect(result).resolves.toBeUndefined()
        expect(visible).toBe('New')
      } else {
        await expect(result).resolves.toBeUndefined()
        expect(visible).toBe('New')
      }
    }

    await runScenario(false)
    await runScenario(true)
  })

  it('reports an attributable declarative prefix as uncertain without restoring it', async () => {
    const textRange = { text: 'Old', load: vi.fn() }
    const textShape = { id: '2', load: vi.fn(), textFrame: { textRange } }
    const geometryShape = {
      id: '3',
      load: vi.fn(),
      left: 10,
      top: 20,
      width: 200,
      height: 40,
    }
    const slide = {
      id: 's1',
      load: vi.fn(),
      shapes: {
        getItem: vi.fn((id: string) => (id === '2' ? textShape : geometryShape)),
      },
    }
    const slides = {
      getCount: vi.fn(() => ({ value: 1 })),
      getItemAt: vi.fn(() => slide),
    }
    let rejected = false
    const sync = vi.fn().mockImplementation(async () => {
      if (textRange.text === 'New' && geometryShape.left === 30 && !rejected) {
        rejected = true
        geometryShape.left = 10
        throw new Error('host rejected after committed prefix')
      }
    })
    Object.assign(globalThis, {
      Office: {
        context: {
          host: 'PowerPoint',
          requirements: { isSetSupported: vi.fn().mockReturnValue(true) },
        },
      },
      PowerPoint: {
        run: (callback: (context: unknown) => unknown) =>
          callback({ presentation: { slides }, sync }),
      },
    })

    await expect(
      new BrowserPowerPointAdapter().executeDeclarative([
        { op: 'set_shape_text', slide_index: 0, shape_id: '2', text: 'New' },
        {
          op: 'set_shape_geometry',
          slide_index: 0,
          shape_id: '3',
          left: 30,
          top: 20,
          width: 200,
          height: 40,
        },
      ]),
    ).rejects.toThrow('office_state_uncertain')
    expect(textRange.text).toBe('New')
    expect(geometryShape.left).toBe(10)
  })

  it('does not restore a declarative batch over a concurrent third state', async () => {
    const textRange = { text: 'Old', load: vi.fn() }
    const shape = { id: '2', load: vi.fn(), textFrame: { textRange } }
    const slide = {
      id: 's1',
      load: vi.fn(),
      shapes: { getItem: vi.fn(() => shape) },
    }
    const slides = {
      getCount: vi.fn(() => ({ value: 1 })),
      getItemAt: vi.fn(() => slide),
    }
    let rejected = false
    const sync = vi.fn().mockImplementation(async () => {
      if (textRange.text === 'New' && !rejected) {
        rejected = true
        textRange.text = 'User edit'
        throw new Error('host rejected after concurrent edit')
      }
    })
    Object.assign(globalThis, {
      Office: {
        context: {
          host: 'PowerPoint',
          requirements: { isSetSupported: vi.fn().mockReturnValue(true) },
        },
      },
      PowerPoint: {
        run: (callback: (context: unknown) => unknown) =>
          callback({ presentation: { slides }, sync }),
      },
    })

    await expect(
      new BrowserPowerPointAdapter().executeDeclarative([
        { op: 'set_shape_text', slide_index: 0, shape_id: '2', text: 'New' },
      ]),
    ).rejects.toThrow('office_concurrent_change')
    expect(textRange.text).toBe('User edit')
  })

  it('rejects repeated writes to the same shape property before Office dispatch', async () => {
    const run = vi.fn()
    Object.assign(globalThis, {
      Office: {
        context: {
          host: 'PowerPoint',
          requirements: { isSetSupported: vi.fn().mockReturnValue(true) },
        },
      },
      PowerPoint: { run },
    })

    await expect(
      new BrowserPowerPointAdapter().executeDeclarative([
        { op: 'set_shape_text', slide_index: 0, shape_id: '2', text: 'First' },
        { op: 'set_shape_text', slide_index: 0, shape_id: '2', text: 'Second' },
      ]),
    ).rejects.toThrow('invalid_tool_input')
    expect(run).not.toHaveBeenCalled()
  })

  it.each([
    { mode: 'committed', expected: 'copy' },
    { mode: 'cancelled', expected: 'cancelled' },
    { mode: 'third-state', expected: 'office_concurrent_change' },
  ])('reconciles duplicate-slide sync rejection in $mode state', async ({ mode, expected }) => {
    const packageZip = new JSZip()
    packageZip.file('ppt/slides/slide1.xml', '<p:sld xmlns:p="urn:p"/>')
    const slidePackage = await packageZip.generateAsync({ type: 'base64' })
    const controller = new AbortController()
    const shape = (text: string) => ({
      id: '2',
      name: 'Title',
      type: 'TextBox',
      left: 10,
      top: 20,
      width: 200,
      height: 40,
      textFrame: { hasText: true, load: vi.fn(), textRange: { text, load: vi.fn() } },
    })
    const slides = {
      items: [] as Array<Record<string, unknown>>,
      getCount: vi.fn(() => ({ value: slides.items.length })),
      getItemAt: vi.fn((index: number) => slides.items[index]),
    }
    const source = {
      id: 'source',
      load: vi.fn(),
      shapes: { load: vi.fn(), items: [shape('Stable')] },
      exportAsBase64: vi.fn(() => ({ value: slidePackage })),
    }
    slides.items.push(source)
    const insertSlidesFromBase64 = vi.fn(() => {
      const inserted = {
        id: 'copy',
        load: vi.fn(),
        shapes: {
          load: vi.fn(),
          items: [shape(mode === 'third-state' ? 'User slide' : 'Stable')],
        },
        exportAsBase64: vi.fn(() => ({ value: slidePackage })),
        delete: vi.fn(() => {
          const index = slides.items.indexOf(inserted)
          if (index >= 0) slides.items.splice(index, 1)
        }),
      }
      slides.items.splice(1, 0, inserted)
    })
    let rejected = false
    const sync = vi.fn().mockImplementation(async () => {
      if (slides.items.length === 2 && !rejected) {
        rejected = true
        if (mode === 'cancelled') controller.abort()
        throw new Error('host rejected after insert')
      }
    })
    Object.assign(globalThis, {
      Office: {
        context: {
          host: 'PowerPoint',
          requirements: { isSetSupported: vi.fn().mockReturnValue(true) },
        },
      },
      PowerPoint: {
        run: (callback: (context: unknown) => unknown) =>
          callback({ presentation: { slides, insertSlidesFromBase64 }, sync }),
      },
    })

    const result = new BrowserPowerPointAdapter().duplicateSlide(0, controller.signal)
    if (mode === 'committed' || mode === 'cancelled')
      await expect(result).resolves.toEqual({ slideId: 'copy' })
    else await expect(result).rejects.toThrow(expected)
    expect(slides.items).toHaveLength(2)
  })

  it('waits for duplicate collection convergence and proves package ownership after sync rejection', async () => {
    const packageZip = new JSZip()
    packageZip.file('ppt/slides/slide1.xml', '<p:sld xmlns:p="urn:p"><p:cSld/></p:sld>')
    const base64 = await packageZip.generateAsync({ type: 'base64' })
    const source = { id: 'source', load: vi.fn(), exportAsBase64: vi.fn(() => ({ value: base64 })) }
    const slides = {
      getCount: vi.fn(() => ({ value: 2 })),
      getItemAt: vi.fn(() => source),
    }
    const insertSlidesFromBase64 = vi.fn()
    let rejected = false
    const sync = vi.fn().mockImplementation(async () => {
      if (insertSlidesFromBase64.mock.calls.length > 0 && !rejected) {
        rejected = true
        throw new Error('host rejected after committed insert')
      }
    })
    Object.assign(globalThis, {
      Office: {
        context: {
          host: 'PowerPoint',
          requirements: { isSetSupported: vi.fn().mockReturnValue(true) },
        },
      },
      PowerPoint: {
        run: (callback: (context: unknown) => unknown) =>
          callback({ presentation: { slides, insertSlidesFromBase64 }, sync }),
      },
    })
    const subject = new BrowserPowerPointAdapter()
    vi.spyOn(subject, 'snapshotSlide')
      .mockResolvedValueOnce({ slideId: 'source', fingerprint: 'source:semantic' })
      .mockResolvedValueOnce({ slideId: 'existing', fingerprint: 'existing:other' })
      .mockResolvedValueOnce({ slideId: 'source', fingerprint: 'source:semantic' })
      .mockResolvedValueOnce({ slideId: 'source', fingerprint: 'source:semantic' })
      .mockResolvedValueOnce({ slideId: 'existing', fingerprint: 'existing:other' })
      .mockResolvedValueOnce({ slideId: 'existing', fingerprint: 'existing:other' })
      .mockResolvedValueOnce({ slideId: 'copy', fingerprint: 'copy:semantic' })
      .mockResolvedValue({ slideId: 'source', fingerprint: 'source:semantic' })
    vi.spyOn(subject, 'exportSlidePackage')
      .mockResolvedValueOnce({ slideId: 'source', base64, fingerprint: 'volatile-source' })
      .mockResolvedValue({ slideId: 'copy', base64, fingerprint: 'volatile-copy' })

    await expect(subject.duplicateSlide(0)).resolves.toEqual({ slideId: 'copy' })
    expect(subject.snapshotSlide).toHaveBeenCalledTimes(8)
    expect(subject.exportSlidePackage).toHaveBeenCalledWith(1)
  })

  it('revalidates the duplicate source immediately before insertion', async () => {
    const packageZip = new JSZip()
    packageZip.file('ppt/slides/slide1.xml', '<p:sld xmlns:p="urn:p"/>')
    const base64 = await packageZip.generateAsync({ type: 'base64' })
    const source = { id: 'source', load: vi.fn() }
    const slides = {
      getCount: vi.fn(() => ({ value: 1 })),
      getItemAt: vi.fn(() => source),
    }
    const insertSlidesFromBase64 = vi.fn()
    Object.assign(globalThis, {
      Office: {
        context: {
          host: 'PowerPoint',
          requirements: { isSetSupported: vi.fn().mockReturnValue(true) },
        },
      },
      PowerPoint: {
        run: (callback: (context: unknown) => unknown) =>
          callback({ presentation: { slides, insertSlidesFromBase64 }, sync: vi.fn() }),
      },
    })
    const subject = new BrowserPowerPointAdapter()
    vi.spyOn(subject, 'snapshotSlide')
      .mockResolvedValueOnce({ slideId: 'source', fingerprint: 'source:before' })
      .mockRejectedValueOnce(new Error('invalid_tool_input'))
      .mockResolvedValue({ slideId: 'source', fingerprint: 'source:user-edit' })
    vi.spyOn(subject, 'exportSlidePackage').mockResolvedValue({
      slideId: 'source',
      base64,
      fingerprint: 'volatile',
    })

    await expect(subject.duplicateSlide(0)).rejects.toThrow('office_concurrent_change')
    expect(insertSlidesFromBase64).not.toHaveBeenCalled()
  })

  it('removes an owned duplicate when the source changes after insertion', async () => {
    const packageZip = new JSZip()
    packageZip.file('ppt/slides/slide1.xml', '<p:sld xmlns:p="urn:p"/>')
    const slidePackage = await packageZip.generateAsync({ type: 'base64' })
    const shape = (text: string) => ({
      id: '2',
      name: 'Title',
      type: 'TextBox',
      left: 10,
      top: 20,
      width: 200,
      height: 40,
      textFrame: { hasText: true, load: vi.fn(), textRange: { text, load: vi.fn() } },
    })
    const sourceShape = shape('Stable')
    const slides = {
      items: [] as Array<Record<string, unknown>>,
      getCount: vi.fn(() => ({ value: slides.items.length })),
      getItemAt: vi.fn((index: number) => slides.items[index]),
    }
    const source = {
      id: 'source',
      load: vi.fn(),
      shapes: { load: vi.fn(), items: [sourceShape] },
      exportAsBase64: vi.fn(() => ({ value: slidePackage })),
    }
    slides.items.push(source)
    const insertSlidesFromBase64 = vi.fn(() => {
      const copy = {
        id: 'copy',
        load: vi.fn(),
        shapes: { load: vi.fn(), items: [shape('Stable')] },
        exportAsBase64: vi.fn(() => ({ value: slidePackage })),
        delete: vi.fn(() => slides.items.splice(slides.items.indexOf(copy), 1)),
      }
      slides.items.splice(1, 0, copy)
      ;(sourceShape.textFrame as { textRange: { text: string } }).textRange.text = 'User edit'
    })
    Object.assign(globalThis, {
      Office: {
        context: {
          host: 'PowerPoint',
          requirements: { isSetSupported: vi.fn().mockReturnValue(true) },
        },
      },
      PowerPoint: {
        run: (callback: (context: unknown) => unknown) =>
          callback({ presentation: { slides, insertSlidesFromBase64 }, sync: vi.fn() }),
      },
    })

    await expect(new BrowserPowerPointAdapter().duplicateSlide(0)).rejects.toThrow(
      'office_concurrent_change',
    )
    expect(slides.items).toHaveLength(2)
    expect(slides.items[0]).toBe(source)
  })

  it('keeps duplicate validation stable when PowerPoint exports volatile slide packages', async () => {
    const sync = vi.fn().mockResolvedValue(undefined)
    const textRange = { text: 'Stable title', load: vi.fn() }
    const textFrame = { hasText: true, load: vi.fn(), textRange }
    const shape = {
      id: '2',
      name: 'Title 1',
      type: 'Placeholder',
      left: 120,
      top: 88.4,
      width: 720,
      height: 188,
      textFrame,
    }
    let exportSequence = 0
    const slide = {
      id: 's1',
      load: vi.fn(),
      shapes: { load: vi.fn(), items: [shape] },
      exportAsBase64: vi.fn(() => ({ value: `volatile-package-${exportSequence++}` })),
    }
    const slides = {
      getCount: vi.fn(() => ({ value: 1 })),
      getItemAt: vi.fn(() => slide),
    }
    Object.assign(globalThis, {
      Office: {
        context: {
          host: 'PowerPoint',
          requirements: { isSetSupported: vi.fn().mockReturnValue(true) },
        },
      },
      PowerPoint: {
        run: (callback: (context: unknown) => unknown) =>
          callback({ presentation: { slides }, sync }),
      },
    })

    const subject = new BrowserPowerPointAdapter()
    const first = await subject.snapshotSlide(0)
    const second = await subject.snapshotSlide(0)

    expect(second).toEqual(first)
    textRange.text = 'Changed title'
    await expect(subject.snapshotSlide(0)).resolves.not.toEqual(first)
  })

  describe('mixed-shape snapshots', () => {
    function fixture(safeFrames: boolean) {
      let pendingTextError: Error | undefined
      let requestedNullFrame = false
      let loadedNullFrame = false
      const font = {
        load: vi.fn(),
        color: '#FFFFFF',
        name: 'Aptos',
        size: 32,
        bold: true,
        italic: false,
      }
      const textRange = { text: 'Readable title', font, load: vi.fn() }
      const textFrame = { isNullObject: false, hasText: true, load: vi.fn(), textRange }
      const nullFrame = {
        get isNullObject() {
          if (!loadedNullFrame) throw new Error('PropertyNotLoaded')
          return true
        },
        load: vi.fn(() => {
          throw new Error('null_text_frame_load')
        }),
        get hasText(): never {
          throw new Error('null_text_frame_access')
        },
      }
      const unsupportedFrame = () => {
        throw new Error('office_api_unsupported')
      }
      const title = {
        id: 'title',
        name: 'Title',
        type: 'TextBox',
        left: 40,
        top: 40,
        width: 600,
        height: 80,
        textFrame,
        getTextFrameOrNullObject: safeFrames ? () => textFrame : unsupportedFrame,
      }
      const image = {
        id: 'image',
        name: 'Cover image',
        type: 'Image',
        left: 0,
        top: 0,
        width: 960,
        height: 540,
        get textFrame(): never {
          throw Object.assign(new Error('InvalidArgument'), {
            code: 'InvalidArgument',
            debugInfo: { errorLocation: 'Shape.textFrame' },
          })
        },
        getTextFrameOrNullObject: safeFrames
          ? () => {
              requestedNullFrame = true
              return nullFrame
            }
          : unsupportedFrame,
      }
      const items: Array<typeof title | typeof image> = [title, image]
      const slide = { id: 's1', load: vi.fn(), shapes: { load: vi.fn(), items } }
      const context = {
        presentation: {
          slides: { getCount: () => ({ value: 1 }), getItemAt: () => slide },
        },
        sync: vi.fn(async () => {
          if (pendingTextError) throw pendingTextError
          if (requestedNullFrame) loadedNullFrame = true
        }),
      }
      Object.assign(globalThis, {
        Office: {
          context: {
            host: 'PowerPoint',
            platform: 'Mac',
            requirements: {
              isSetSupported: (_name: string, version: string) => version !== '1.10' || safeFrames,
            },
          },
        },
        PowerPoint: { run: (callback: (value: typeof context) => unknown) => callback(context) },
      })
      return {
        subject: new BrowserPowerPointAdapter(),
        textRange,
        font,
        image,
        items,
        failTextRead(error: Error) {
          textFrame.load.mockImplementation(() => {
            pendingTextError = error
          })
        },
      }
    }

    it.each([false, true])(
      'snapshots text and images without unsafe textFrame access (API 1.10: %s)',
      async (safeFrames) => {
        const { subject, textRange, font, image, items } = fixture(safeFrames)
        const before = await subject.snapshotSlide(0)
        await expect(subject.snapshotSlide(0)).resolves.toEqual(before)
        textRange.text = 'Changed title'
        const changedText = await subject.snapshotSlide(0)
        expect(changedText.fingerprint).not.toBe(before.fingerprint)
        font.size = 40
        const changedStyle = await subject.snapshotSlide(0)
        expect(changedStyle.fingerprint).not.toBe(changedText.fingerprint)
        image.left = 20
        const movedImage = await subject.snapshotSlide(0)
        expect(movedImage.fingerprint).not.toBe(changedStyle.fingerprint)
        items.pop()
        const removedImage = await subject.snapshotSlide(0)
        expect(removedImage.fingerprint).not.toBe(movedImage.fingerprint)
      },
    )

    it('does not skip an unknown legacy shape type when its text read fails', async () => {
      const { subject, image } = fixture(false)
      image.type = 'FutureShape'
      await expect(subject.snapshotSlide(0)).rejects.toMatchObject({
        code: 'InvalidArgument',
        debugInfo: { errorLocation: 'Shape.textFrame' },
      })
    })

    it.each([false, true])(
      'does not suppress actual text read failures (API 1.10: %s)',
      async (safeFrames) => {
        const { subject, items, failTextRead } = fixture(safeFrames)
        items.pop()
        const hostError = Object.assign(new Error('GeneralException'), { code: 'GeneralException' })
        failTextRead(hostError)
        await expect(subject.snapshotSlide(0)).rejects.toBe(hostError)
      },
    )
  })

  it('rejects empty or oversized duplicate exports before insertion', async () => {
    const sync = vi.fn().mockResolvedValue(undefined)
    const slide = {
      id: 's1',
      load: vi.fn(),
      shapes: { load: vi.fn(), items: [] },
      exportAsBase64: vi.fn(() => ({ value: '' })),
    }
    const slides = {
      items: [slide],
      getCount: vi.fn(() => ({ value: 1 })),
      getItemAt: vi.fn(() => slide),
    }
    const insertSlidesFromBase64 = vi.fn()
    Object.assign(globalThis, {
      Office: {
        context: {
          host: 'PowerPoint',
          requirements: { isSetSupported: vi.fn().mockReturnValue(true) },
        },
      },
      PowerPoint: {
        run: (callback: (context: unknown) => unknown) =>
          callback({ presentation: { slides, insertSlidesFromBase64 }, sync }),
      },
    })
    const subject = new BrowserPowerPointAdapter()
    await expect(subject.duplicateSlide(0)).rejects.toThrow('office_write_failed')
    slide.exportAsBase64.mockReturnValueOnce({ value: 'x'.repeat(8 * 1024 * 1024 + 1) })
    await expect(subject.duplicateSlide(0)).rejects.toThrow('office_write_failed')
    const controller = new AbortController()
    sync.mockClear()
    sync.mockImplementation(async () => {
      if (sync.mock.calls.length === 3) controller.abort()
    })
    slide.exportAsBase64.mockReturnValueOnce({ value: 'ppt' })
    await expect(subject.duplicateSlide(0, controller.signal)).rejects.toThrow('cancelled')
    expect(insertSlidesFromBase64).not.toHaveBeenCalled()
  })

  it('cancels package replacement before queuing its irreversible insert/delete batch', async () => {
    const controller = new AbortController()
    const sync = vi.fn().mockImplementation(async () => {
      if (sync.mock.calls.length === 2) controller.abort()
    })
    const remove = vi.fn()
    const slide = { id: 's1', load: vi.fn(), delete: remove }
    const slides = {
      getCount: vi.fn(() => ({ value: 1 })),
      getItemAt: vi.fn(() => slide),
    }
    const insertSlidesFromBase64 = vi.fn()
    Object.assign(globalThis, {
      Office: {
        context: {
          host: 'PowerPoint',
          requirements: { isSetSupported: vi.fn().mockReturnValue(true) },
        },
      },
      PowerPoint: {
        run: (callback: (context: unknown) => unknown) =>
          callback({ presentation: { slides, insertSlidesFromBase64 }, sync }),
      },
    })
    await expect(
      new BrowserPowerPointAdapter().replaceSlidePackage(
        0,
        'ppt',
        false,
        undefined,
        controller.signal,
      ),
    ).rejects.toThrow('cancelled')
    expect(insertSlidesFromBase64).not.toHaveBeenCalled()
    expect(remove).not.toHaveBeenCalled()
  })

  it('never restores a package replacement over an unowned third state', async () => {
    const originalZip = new JSZip()
    originalZip.file('ppt/slides/slide1.xml', '<p:sld xmlns:p="urn:p"><p:cSld/></p:sld>')
    const originalPackage = await originalZip.generateAsync({ type: 'base64' })
    const expected = await editPowerPointPackage(originalPackage, 'slide', [
      {
        path: 'ppt/slides/slide1.xml',
        xml: '<p:sld xmlns:p="urn:p"><p:cSld><p:sp/></p:cSld></p:sld>',
      },
    ])
    const thirdZip = new JSZip()
    thirdZip.file('ppt/slides/slide1.xml', '<p:sld xmlns:p="urn:p"><p:user/></p:sld>')
    const thirdPackage = await thirdZip.generateAsync({ type: 'base64' })
    const slides = {
      items: [] as Array<Record<string, unknown>>,
      getCount: vi.fn(() => ({ value: slides.items.length })),
      getItemAt: vi.fn((index: number) => slides.items[index]),
    }
    const makeSlide = (id: string, exported: string) => {
      const item = {
        id,
        load: vi.fn(),
        exportAsBase64: vi.fn(() => ({ value: exported })),
        delete: vi.fn(() => {
          slides.items = slides.items.filter((slide) => slide !== item)
        }),
      }
      return item
    }
    const original = makeSlide('source', originalPackage)
    const third = makeSlide('user-slide', thirdPackage)
    slides.items = [original]
    let batchQueued = false
    let failed = false
    const insertSlidesFromBase64 = vi.fn((value: string) => {
      batchQueued = true
      slides.items.splice(0, 0, makeSlide('imported', value))
    })
    const sync = vi.fn().mockImplementation(async () => {
      if (batchQueued && !failed) {
        failed = true
        slides.items = [third]
        throw new Error('host rejected after third-party replacement')
      }
    })
    Object.assign(globalThis, {
      Office: {
        context: {
          host: 'PowerPoint',
          requirements: { isSetSupported: vi.fn().mockReturnValue(true) },
        },
      },
      PowerPoint: {
        run: (callback: (context: unknown) => unknown) =>
          callback({ presentation: { slides, insertSlidesFromBase64 }, sync }),
      },
    })

    await expect(
      new BrowserPowerPointAdapter().replaceSlidePackage(0, expected.base64, false, expected),
    ).rejects.toThrow(/office_(concurrent_change|state_uncertain)/)
    expect(slides.items).toEqual([third])
    expect(insertSlidesFromBase64).toHaveBeenCalledOnce()
  })

  it.each([
    'sync-failure',
    'ignored-package-import',
    'ignored-layout-recovery',
    'wrong-restored-slide',
  ])('never performs an unowned package or layout restore after %s', async (mode) => {
    const packageZip = new JSZip()
    packageZip.file('ppt/slides/slide1.xml', '<p:sld xmlns:p="urn:p"/>')
    const originalPackage = await packageZip.generateAsync({ type: 'base64' })
    const expected = await editPowerPointPackage(originalPackage, 'slide', [
      { path: 'ppt/slides/slide1.xml', xml: '<p:sld xmlns:p="urn:p"><p:cSld/></p:sld>' },
    ])
    const oldLayout = { id: 'old-layout', name: '', load: vi.fn() }
    const newLayout = { id: 'new-layout', name: '', load: vi.fn() }
    const wrongLayout = { id: 'wrong-layout', name: '', load: vi.fn() }
    const slides: {
      items: any[]
      getCount: ReturnType<typeof vi.fn>
      getItemAt: ReturnType<typeof vi.fn>
      load: ReturnType<typeof vi.fn>
    } = {
      items: [],
      getCount: vi.fn(() => ({ value: slides.items.length })),
      getItemAt: vi.fn((index: number) => slides.items[index]),
      load: vi.fn(),
    }
    const makeSlide = (id: string, layout: any, exported = 'original') => {
      const item: any = {
        id,
        layout,
        load: vi.fn(),
        exportAsBase64: vi.fn(() => ({ value: exported })),
        applyLayout: vi.fn((next) => {
          item.layout = next
        }),
      }
      item.delete = vi.fn(() => {
        slides.items = slides.items.filter((slide) => slide !== item)
      })
      return item
    }
    const original = makeSlide('s1', oldLayout, originalPackage)
    const sibling = makeSlide('s-other', oldLayout)
    slides.items = [original, sibling]
    const oldMaster = { layouts: { items: [oldLayout], load: vi.fn() } }
    const newMaster = { layouts: { items: [newLayout], load: vi.fn() } }
    const masters = { items: [oldMaster, newMaster], load: vi.fn() }
    let propagationStarted = false
    sibling.applyLayout.mockImplementation((next: unknown) => {
      if (mode === 'ignored-layout-recovery') {
        if (next === newLayout) sibling.layout = wrongLayout
      } else {
        sibling.layout = next
      }
      propagationStarted = true
    })
    let failed = false
    const sync = vi.fn().mockImplementation(async () => {
      if (mode !== 'ignored-layout-recovery' && propagationStarted && !failed) {
        failed = true
        throw new Error('host failure')
      }
    })
    const insertSlidesFromBase64 = vi.fn(
      (base64: string, _options: { formatting: string; targetSlideId?: string }) => {
        const inserted = makeSlide(
          base64 === expected.base64 ? 's2' : 's1-restored',
          base64 === expected.base64 ? newLayout : oldLayout,
          mode === 'ignored-package-import' && base64 === expected.base64
            ? originalPackage
            : mode === 'wrong-restored-slide' && base64 === originalPackage
              ? expected.base64
              : base64,
        )
        slides.items.splice(0, 0, inserted)
      },
    )
    Object.assign(globalThis, {
      Office: {
        context: {
          host: 'PowerPoint',
          requirements: { isSetSupported: vi.fn().mockReturnValue(true) },
        },
      },
      PowerPoint: {
        run: (callback: (context: unknown) => unknown) =>
          callback({
            presentation: { slides, slideMasters: masters, insertSlidesFromBase64 },
            sync,
          }),
      },
    })
    const failure = await new BrowserPowerPointAdapter()
      .replaceSlidePackage(0, expected.base64, true, expected)
      .catch((error: unknown) => error)
    expect(failure).toMatchObject({ message: 'office_state_uncertain' })
    if (mode === 'ignored-package-import')
      expect(failure).toMatchObject({
        debugInfo: { errorLocation: 'PowerPoint.replaceSlidePackage.packageImportVerify' },
      })
    if (mode === 'ignored-layout-recovery')
      expect(failure).toMatchObject({
        debugInfo: { errorLocation: 'PowerPoint.replaceSlidePackage.layoutApplyVerify' },
      })
    expect(insertSlidesFromBase64).toHaveBeenCalledOnce()
    expect(insertSlidesFromBase64.mock.calls[0]?.[1]).toMatchObject({
      formatting: 'KeepSourceFormatting',
    })
    expect(slides.items).toHaveLength(2)
    expect(slides.items[0].id).toBe('s2')
  })

  it('treats a text-only change as stale even when slide geometry is unchanged', async () => {
    const fake = adapter()
    const proposals = createStructuredProposalController()
    const skill = createPowerPointSkill({ adapter: fake, proposals })
    await skill.executeTool(
      call('edit_slide_text', { slide_index: 0, shape_id: '2', text: 'Replacement' }),
    )
    ;(fake.readSlideText as ReturnType<typeof vi.fn>).mockResolvedValueOnce({
      slideId: 'slide-1',
      shapeId: '2',
      text: 'Changed elsewhere',
      paragraphs: ['Changed elsewhere'],
    })
    await expect(proposals.confirm(proposals.pending()!.id)).rejects.toThrow('proposal_stale')
    expect(fake.editSlideText).not.toHaveBeenCalled()
  })
})
